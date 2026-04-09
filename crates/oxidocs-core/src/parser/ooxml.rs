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
        // Parse theme first — needed for font resolution in styles
        let theme = match self.read_part("word/theme/theme1.xml") {
            Ok(xml) => parse_theme(&xml),
            Err(_) => ThemeColors::default(),
        };
        let styles = self.parse_styles_with_theme(&theme)?;
        let numbering = self.parse_numbering()?;
        let ctx = self.build_context_with_theme(numbering, theme, &styles)?;
        let sections = self.parse_document_xml(&ctx, &styles)?;
        let metadata = self.parse_metadata();
        let adjust_line_height_in_table = self.parse_adjust_line_height_in_table();
        let default_tab_stop = self.parse_default_tab_stop();
        let compat_mode = self.parse_compat_mode();
        let compress_punctuation = self.parse_compress_punctuation();

        let mut pages = Vec::new();
        let mut page_index = 0usize;
        // OOXML: sections without explicit header/footer inherit from previous section
        let mut prev_header_refs: Vec<HdrFtrRef> = Vec::new();
        let mut prev_footer_refs: Vec<HdrFtrRef> = Vec::new();
        for mut section in sections {
            let effective_header_refs = if section.properties.header_refs.is_empty() {
                &prev_header_refs
            } else {
                &section.properties.header_refs
            };
            let effective_footer_refs = if section.properties.footer_refs.is_empty() {
                &prev_footer_refs
            } else {
                &section.properties.footer_refs
            };
            // Determine which header/footer type to use
            // First page of a section with title_pg uses "first" type
            let hdr_type = if section.properties.title_pg && page_index == 0 { "first" } else { "default" };
            let use_headers: Vec<HdrFtrRef> = effective_header_refs.iter()
                .filter(|r| r.ref_type == hdr_type)
                .cloned()
                .collect();
            let use_footers: Vec<HdrFtrRef> = effective_footer_refs.iter()
                .filter(|r| r.ref_type == hdr_type)
                .cloned()
                .collect();
            // Fall back: if no matching type, try any available reference
            let header = if use_headers.is_empty() {
                let fallback: Vec<HdrFtrRef> = effective_header_refs.iter()
                    .filter(|r| r.ref_type == "default").cloned().collect();
                if fallback.is_empty() {
                    self.parse_header_footer_blocks(effective_header_refs, &ctx, &styles)
                } else {
                    self.parse_header_footer_blocks(&fallback, &ctx, &styles)
                }
            } else {
                self.parse_header_footer_blocks(&use_headers, &ctx, &styles)
            };
            let footer = if use_footers.is_empty() {
                let fallback: Vec<HdrFtrRef> = effective_footer_refs.iter()
                    .filter(|r| r.ref_type == "default").cloned().collect();
                if fallback.is_empty() {
                    self.parse_header_footer_blocks(effective_footer_refs, &ctx, &styles)
                } else {
                    self.parse_header_footer_blocks(&fallback, &ctx, &styles)
                }
            } else {
                self.parse_header_footer_blocks(&use_footers, &ctx, &styles)
            };

            // Collect referenced footnotes and endnotes for this section
            let mut footnotes_list = Vec::new();
            let mut endnotes_list = Vec::new();
            collect_note_refs(&section.blocks, &ctx, &mut footnotes_list, &mut endnotes_list);
            footnotes_list.sort_by_key(|f| f.number);
            endnotes_list.sort_by_key(|f| f.number);

            // Round 29 (2026-04-08): renumber footnote/endnote references
            // sequentially per section. The OOXML <w:footnoteReference w:id="N"/>
            // values are NOT necessarily 1,2,3... — they can start from 2 (id=1
            // is reserved for the separator) and have arbitrary gaps. Word
            // displays them as 1,2,3..., so we need to remap. The parser had
            // set Run.text = "[id]" which is wrong; rewrite it here to "[seq]".
            let mut fn_id_to_seq: std::collections::HashMap<u32, u32> = std::collections::HashMap::new();
            for (i, f) in footnotes_list.iter().enumerate() {
                fn_id_to_seq.insert(f.number, (i as u32) + 1);
            }
            let mut en_id_to_seq: std::collections::HashMap<u32, u32> = std::collections::HashMap::new();
            for (i, f) in endnotes_list.iter().enumerate() {
                en_id_to_seq.insert(f.number, (i as u32) + 1);
            }
            renumber_note_refs(&mut section.blocks, &fn_id_to_seq, &en_id_to_seq);

            // Continuous section: merge into previous page instead of creating a new one
            if section.properties.section_type.as_deref() == Some("continuous") && !pages.is_empty() {
                let last: &mut Page = pages.last_mut().unwrap();
                last.blocks.extend(section.blocks);
                last.floating_images.extend(section.floating_images);
                last.text_boxes.extend(section.text_boxes);
                last.shapes.extend(section.shapes);
                last.footnotes.extend(footnotes_list);
                last.endnotes.extend(endnotes_list);
                if section.properties.columns.is_some() {
                    last.columns = section.properties.columns;
                }
            } else {
                pages.push(Page {
                    blocks: section.blocks,
                    size: section.properties.page_size,
                    margin: section.properties.margin,
                    grid_line_pitch: section.properties.grid_line_pitch,
                    grid_char_pitch: section.properties.grid_char_pitch,
                    doc_grid_no_type: section.properties.doc_grid_no_type,
                    header,
                    footer,
                    footnotes: footnotes_list,
                    endnotes: endnotes_list,
                    floating_images: section.floating_images,
                    text_boxes: section.text_boxes,
                    shapes: section.shapes,
                    columns: section.properties.columns,
                    header_distance: section.properties.header_distance,
                    footer_distance: section.properties.footer_distance,
                    page_number_format: section.properties.page_number_format,
                    page_number_start: section.properties.page_number_start,
                    page_borders: section.properties.page_borders,
                });
            }
            // Update previous refs for inheritance
            if !section.properties.header_refs.is_empty() {
                prev_header_refs = section.properties.header_refs;
            }
            if !section.properties.footer_refs.is_empty() {
                prev_footer_refs = section.properties.footer_refs;
            }
            page_index += 1;
        }

        // Post-process: fix charGrid pitch using actual default font size.
        // parse_section_properties uses 10.5pt placeholder; correct with Normal style font size.
        let default_font_size = styles.styles.get("Normal")
            .or_else(|| styles.styles.get("a"))
            .and_then(|s| s.paragraph.default_run_style.as_ref())
            .and_then(|rs| rs.font_size)
            .or_else(|| styles.doc_default_run_style.as_ref().and_then(|rs| rs.font_size))
            .unwrap_or(10.5);
        for page in &mut pages {
            if let Some(pitch) = page.grid_char_pitch {
                // Recalculate: the pitch was computed with 10.5pt; redo with actual default size
                let content_w = page.size.width - page.margin.left - page.margin.right;
                if content_w > 0.0 {
                    // Reverse-engineer the charSpace from the stored pitch
                    // pitch = contentW / floor(contentW / raw_pitch)
                    // We need to recompute with correct default_font_size
                    // Original raw_pitch = 10.5 + charSpace_pt; we need default_font_size + charSpace_pt
                    // charSpace_pt = original_raw_pitch - 10.5 = contentW/charsLine_old - 10.5... complex
                    // Simpler: just recalculate from scratch using the page margins
                    // Since we can't easily get charSpace here, use a different approach:
                    // If the pitch was based on 10.5 but should be based on default_font_size,
                    // and both use the same formula, we can just redo it.
                    // But we lost the charSpace value. Store it in SectionProperties.
                    // For now: if default != 10.5, redo with approximate raw_pitch
                    if (default_font_size - 10.5).abs() > 0.01 {
                        // Approximate: old raw_pitch = old_pitch / ceil (which was just floor(cw/raw)/cw)
                        // Actually, we stored pitch = cw/floor(cw/10.5). Just redo:
                        let old_chars_line = (content_w / pitch).round();
                        // The charSpace contribution: old raw = 10.5 + cs, new raw = default + cs
                        // cs = content_w/old_chars_line - content_w/floor(content_w/(10.5)) ... no
                        // Simplest: just use default_font_size as raw_pitch (no charSpace contribution)
                        // This works when charSpace is absent (most common case for this bug)
                        let raw_pitch = default_font_size;
                        let chars_line = (content_w / raw_pitch).floor().max(1.0);
                        page.grid_char_pitch = Some(content_w / chars_line);
                    }
                }
            }
        }

        // Collect all comments referenced in the document
        let all_comments: Vec<Comment> = ctx.comments.values().cloned().collect();

        Ok(Document {
            pages,
            styles,
            metadata,
            comments: all_comments,
            adjust_line_height_in_table,
            default_tab_stop,
            compat_mode,
            compress_punctuation,
        })
    }

    /// Parse header or footer XML parts referenced by relationship IDs
    fn parse_header_footer_blocks(
        &mut self,
        refs: &[HdrFtrRef],
        ctx: &ParseContext,
        styles: &StyleSheet,
    ) -> Vec<Block> {
        let mut blocks = Vec::new();
        for hdr_ref in refs {
            // Look up the relationship target path
            let target = ctx._rels.get(&hdr_ref.rel_id)
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

    fn build_context_with_theme(&mut self, numbering: NumberingDefinitions, theme: ThemeColors, styles: &StyleSheet) -> Result<ParseContext, ParseError> {
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
                // Skip external targets (e.g., file:/// URLs)
                if rel.target.starts_with("file:") || rel.target.starts_with("http:") || rel.target.starts_with("https:") {
                    continue;
                }
                // Validate relationship target path against traversal attacks
                let path = match oxi_common::security::sanitize_rel_target("word", &rel.target) {
                    Ok(p) => p,
                    Err(e) => {
                        continue;
                    }
                };
                if let Ok(data) = self.read_binary_part(&path) {
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

        // Parse footnotes (Round 29: pass styles for pStyle inheritance —
        // footnote text style "a8" sets snapToGrid=0, which without
        // inheritance leaves footnote paragraphs grid-snapped to body
        // pitch and causes line wrap to be ~5 chars too narrow.)
        let footnotes = match self.read_part("word/footnotes.xml") {
            Ok(xml) => parse_notes_xml(&xml, styles)?,
            Err(ParseError::MissingPart(_)) => HashMap::new(),
            Err(e) => return Err(e),
        };

        // Parse endnotes
        let endnotes = match self.read_part("word/endnotes.xml") {
            Ok(xml) => parse_notes_xml(&xml, styles)?,
            Err(ParseError::MissingPart(_)) => HashMap::new(),
            Err(e) => return Err(e),
        };

        // Parse comments
        let comments = match self.read_part("word/comments.xml") {
            Ok(xml) => parse_comments_xml(&xml)?,
            Err(ParseError::MissingPart(_)) => HashMap::new(),
            Err(e) => return Err(e),
        };

        // Theme already parsed and passed in
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

    fn parse_styles_with_theme(&mut self, theme: &ThemeColors) -> Result<StyleSheet, ParseError> {
        match self.read_part("word/styles.xml") {
            Ok(xml) => parse_styles(&xml, theme),
            Err(ParseError::MissingPart(_)) => Ok(StyleSheet::default()),
            Err(e) => Err(e),
        }
    }

    fn parse_metadata(&self) -> DocumentMetadata {
        DocumentMetadata::default()
    }

    /// Parse word/settings.xml for adjustLineHeightInTable compatibility flag.
    /// COM measurement (2026-03-27): Compatibility(12) = False for ALL tested documents,
    /// regardless of whether <w:adjustLineHeightInTable/> is present in XML.
    /// 151 documents tested across compatibilityMode 14 and 15.
    /// Therefore: always return false (= table cells snap to grid like normal paragraphs).
    fn parse_adjust_line_height_in_table(&mut self) -> bool {
        false
    }

    /// Parse word/settings.xml for compatibilityMode.
    fn parse_compat_mode(&mut self) -> u32 {
        let xml = match self.read_part("word/settings.xml") {
            Ok(x) => x,
            Err(_) => return 15, // default to Word 2013+
        };
        let mut reader = Reader::from_str(&xml);
        loop {
            match reader.read_event() {
                Ok(Event::Empty(e)) => {
                    if local_name(e.name().as_ref()) == "compatSetting" {
                        let mut is_compat_mode = false;
                        let mut val = String::new();
                        for attr in e.attributes().flatten() {
                            let key = local_name(attr.key.as_ref());
                            let v = String::from_utf8_lossy(&attr.value).to_string();
                            if key == "name" && v == "compatibilityMode" { is_compat_mode = true; }
                            if key == "val" { val = v; }
                        }
                        if is_compat_mode {
                            return val.parse().unwrap_or(15);
                        }
                    }
                }
                Ok(Event::Eof) => break,
                Err(_) => break,
                _ => {}
            }
        }
        15
    }

    /// Parse word/settings.xml for w:characterSpacingControl value.
    /// Returns true if value is "compressPunctuation" or "compressPunctuationAndJapaneseKana"
    /// (enables yakumono compression). False for "doNotCompress" or absent.
    fn parse_compress_punctuation(&mut self) -> bool {
        let xml = match self.read_part("word/settings.xml") {
            Ok(x) => x,
            Err(_) => return false,
        };
        let mut reader = Reader::from_str(&xml);
        loop {
            match reader.read_event() {
                Ok(Event::Empty(e)) => {
                    if local_name(e.name().as_ref()) == "characterSpacingControl" {
                        for attr in e.attributes().flatten() {
                            if local_name(attr.key.as_ref()) == "val" {
                                let val = String::from_utf8_lossy(&attr.value);
                                return val == "compressPunctuation"
                                    || val == "compressPunctuationAndJapaneseKana";
                            }
                        }
                    }
                }
                Ok(Event::Eof) => break,
                Err(_) => break,
                _ => {}
            }
        }
        false
    }

    /// Parse word/settings.xml for defaultTabStop value.
    fn parse_default_tab_stop(&mut self) -> Option<f32> {
        let xml = match self.read_part("word/settings.xml") {
            Ok(x) => x,
            Err(_) => return None,
        };
        // Quick search for w:defaultTabStop
        let mut reader = Reader::from_str(&xml);
        let mut buf: Vec<u8> = Vec::new();
        loop {
            match reader.read_event() {
                Ok(Event::Empty(e)) => {
                    if local_name(e.name().as_ref()) == "defaultTabStop" {
                        for attr in e.attributes().flatten() {
                            if local_name(attr.key.as_ref()) == "val" {
                                let val = String::from_utf8_lossy(&attr.value);
                                if let Ok(twips) = val.parse::<f32>() {
                                    return Some(twips / 20.0);
                                }
                            }
                        }
                    }
                }
                Ok(Event::Eof) => break,
                Err(_) => break,
                _ => {}
            }
            buf.clear();
        }
        None
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
    shapes: Vec<Shape>,
}

/// Parse the w:body content of document.xml into sections
fn parse_body(xml: &str, ctx: &ParseContext, styles: &StyleSheet) -> Result<Vec<ParsedSection>, ParseError> {
    let mut reader = Reader::from_str(xml);
    let mut sections: Vec<ParsedSection> = Vec::new();
    let mut current_blocks = Vec::new();
    let mut current_floating_images: Vec<Image> = Vec::new();
    let mut current_text_boxes: Vec<TextBox> = Vec::new();
    let mut current_shapes: Vec<Shape> = Vec::new();
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
                        let pr = parse_paragraph(&mut reader, ctx, styles)?;
                        current_blocks.push(Block::Paragraph(pr.paragraph));
                        // Inline images become separate blocks after the paragraph
                        current_blocks.extend(pr.inline_images);
                        // Set anchor_block_index for floating images
                        let anchor_idx = current_blocks.len().saturating_sub(1);
                        for mut img in pr.floating_images {
                            img.anchor_block_index = anchor_idx;
                            current_floating_images.push(img);
                        }
                        current_shapes.extend(pr.shapes);
                        for mut tb in pr.text_boxes {
                            tb.anchor_block_index = anchor_idx;
                            current_text_boxes.push(tb);
                        }
                        // If this paragraph contained a section break, start a new section
                        if let Some(sp) = pr.sect_pr {
                            sections.push(ParsedSection {
                                blocks: std::mem::take(&mut current_blocks),
                                properties: sp,
                                floating_images: std::mem::take(&mut current_floating_images),
                                text_boxes: std::mem::take(&mut current_text_boxes),
                                shapes: std::mem::take(&mut current_shapes),
                            });
                        }
                    }
                    "tbl" if in_body && depth == 0 => {
                        let table = parse_table(&mut reader, ctx, styles)?;
                        current_blocks.push(Block::Table(table));
                    }
                    "sdt" if in_body && depth == 0 => {
                        // Structured Document Tag — skip sdtPr, process sdtContent
                        let mut sdt_depth = 1u32;
                        let mut in_sdt_content = false;
                        loop {
                            match reader.read_event()? {
                                Event::Start(se) => {
                                    let sl = local_name(se.name().as_ref());
                                    if sl == "sdtContent" {
                                        in_sdt_content = true;
                                    } else if in_sdt_content {
                                        match sl.as_str() {
                                            "p" => {
                                                let pr = parse_paragraph(&mut reader, ctx, styles)?;
                                                current_blocks.push(Block::Paragraph(pr.paragraph));
                                                current_blocks.extend(pr.inline_images);
                                                let anchor_idx2 = current_blocks.len().saturating_sub(1);
                                                for mut img in pr.floating_images {
                                                    img.anchor_block_index = anchor_idx2;
                                                    current_floating_images.push(img);
                                                }
                                                current_shapes.extend(pr.shapes);
                                                for mut tb in pr.text_boxes {
                                                    tb.anchor_block_index = anchor_idx2;
                                                    current_text_boxes.push(tb);
                                                }
                                                if let Some(sp) = pr.sect_pr {
                                                    sections.push(ParsedSection {
                                                        blocks: std::mem::take(&mut current_blocks),
                                                        properties: sp,
                                                        floating_images: std::mem::take(&mut current_floating_images),
                                                        text_boxes: std::mem::take(&mut current_text_boxes),
                                                        shapes: std::mem::take(&mut current_shapes),
                                                    });
                                                }
                                            }
                                            "tbl" => {
                                                let table = parse_table(&mut reader, ctx, styles)?;
                                                current_blocks.push(Block::Table(table));
                                            }
                                            _ => { sdt_depth += 1; }
                                        }
                                    } else {
                                        sdt_depth += 1;
                                    }
                                }
                                Event::End(ee) => {
                                    let sl = local_name(ee.name().as_ref());
                                    if sl == "sdtContent" {
                                        in_sdt_content = false;
                                    } else if sl == "sdt" {
                                        break;
                                    } else if sdt_depth > 0 {
                                        sdt_depth -= 1;
                                    }
                                }
                                Event::Eof => break,
                                _ => {}
                            }
                        }
                    }
                    // mc:AlternateContent at body level (e.g., SmartArt diagrams)
                    "AlternateContent" if in_body && depth == 0 => {
                        let ac = parse_alternate_content(&mut reader, &ctx, &styles)?;
                        if let Some(drawing) = ac {
                            if let Some(image) = drawing.image {
                                if image.position.is_some() {
                                    current_floating_images.push(image);
                                } else {
                                    current_blocks.push(Block::Image(image));
                                }
                            }
                            if let Some(mut shape) = drawing.shape {
                                shape.anchor_block_index = current_blocks.len().saturating_sub(1);
                                current_shapes.push(shape);
                            }
                            if let Some(mut tb) = drawing.text_box {
                                tb.anchor_block_index = current_blocks.len().saturating_sub(1);
                                current_text_boxes.push(tb);
                            }
                        }
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
            Event::Empty(e) => {
                let local = local_name(e.name().as_ref());
                match local.as_str() {
                    // Self-closing <w:p .../> — empty paragraph with no children.
                    // Apply default Normal style + docDefaults, matching parse_paragraph().
                    "p" if in_body && depth == 0 => {
                        let mut style = ParagraphStyle::default();

                        // Apply Normal style (common IDs: "a" for Japanese, "Normal" for English)
                        if let Some(defined) = styles.styles.get("a")
                            .or_else(|| styles.styles.get("Normal"))
                        {
                            let ds = &defined.paragraph;
                            if style.space_before.is_none() { style.space_before = ds.space_before; }
                            if style.space_after.is_none() { style.space_after = ds.space_after; }
                            if style.line_spacing.is_none() {
                                style.line_spacing = ds.line_spacing;
                                style.line_spacing_rule = ds.line_spacing_rule.clone();
                            }
                            if let Some(ref drs) = ds.default_run_style {
                                style.default_run_style = Some(drs.clone());
                            }
                            if ds.keep_next { style.keep_next = true; }
                            if ds.keep_lines { style.keep_lines = true; }
                        }
                        // Apply docDefaults fallback
                        if style.default_run_style.is_none() {
                            style.default_run_style = styles.doc_default_run_style.clone();
                        }
                        if let Some(ref doc_para) = styles.doc_default_para_style {
                            if style.space_before.is_none() { style.space_before = doc_para.space_before; }
                            if style.space_after.is_none() { style.space_after = doc_para.space_after; }
                            if style.line_spacing.is_none() {
                                style.line_spacing = doc_para.line_spacing;
                                style.line_spacing_rule = doc_para.line_spacing_rule.clone();
                                style.line_spacing_from_doc_defaults = true;
                            }
                            if style.indent_left.is_none() { style.indent_left = doc_para.indent_left; }
                            if style.indent_right.is_none() { style.indent_right = doc_para.indent_right; }
                            if style.indent_first_line.is_none() { style.indent_first_line = doc_para.indent_first_line; }
                            // Empty paragraphs never have explicit widowControl
                            style.widow_control = doc_para.widow_control;
                        }

                        current_blocks.push(Block::Paragraph(Paragraph {
                            runs: vec![],
                            style,
                            alignment: Alignment::default(),
                            shapes: vec![],
                        }));
                    }
                    _ => {}
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
        grid_char_pitch: None,
        doc_grid_no_type: false,
        header_refs: Vec::new(),
        footer_refs: Vec::new(),
        columns: None,
        title_pg: false,
        section_type: None,
        page_number_format: None,
        page_number_start: None,
        page_borders: None,
        header_distance: None,
        footer_distance: None,
    });
    sections.push(ParsedSection {
        blocks: current_blocks,
        properties: last_sp,
        floating_images: current_floating_images,
        text_boxes: current_text_boxes,
        shapes: current_shapes,
    });

    // §17.2.2 / Round 10 (2026-04-08, COM-confirmed):
    // Word implicitly creates a single empty body paragraph when the
    // <w:body> contains only <w:sectPr> with no <w:p> or <w:tbl> elements
    // (e.g., header_page_number_01, footer_complex_01). Without this,
    // Oxi produces 0 body paragraphs and any tooling that expects a
    // body paragraph (dml_diff, layout cursor placement) breaks.
    if sections.iter().all(|s| s.blocks.is_empty()) {
        if let Some(sec) = sections.last_mut() {
            sec.blocks.push(Block::Paragraph(Paragraph {
                runs: Vec::new(),
                style: ParagraphStyle::default(),
                alignment: Alignment::Left,
                shapes: Vec::new(),
            }));
        }
    }

    Ok(sections)
}

/// Parse a w:p element (paragraph).
/// Returns (Paragraph, optional SectionProperties if this paragraph ends a section).
/// Parsed paragraph plus any floating elements found inside it
struct ParagraphResult {
    paragraph: Paragraph,
    sect_pr: Option<SectionProperties>,
    shapes: Vec<Shape>,
    text_boxes: Vec<TextBox>,
    inline_images: Vec<Block>,
    floating_images: Vec<Image>,
}

fn parse_paragraph(reader: &mut Reader<&[u8]>, ctx: &ParseContext, styles: &StyleSheet) -> Result<ParagraphResult, ParseError> {
    let mut runs = Vec::new();
    let mut images = Vec::new();
    let mut found_shapes: Vec<Shape> = Vec::new();
    let mut found_text_boxes: Vec<TextBox> = Vec::new();
    let mut style = ParagraphStyle::default();
    let mut alignment = Alignment::default();
    let mut style_id: Option<String> = None;
    let mut num_pr_ref: Option<NumPrRef> = None;
    let mut para_sect_pr: Option<SectionProperties> = None;
    let mut depth = 0;
    // Field state: tracks fldChar begin/separate/end across runs.
    // Runs between "separate" and "end" contain cached field results
    // that should be suppressed when the field is evaluated (e.g. PAGE).
    let mut field_result_depth: i32 = 0; // >0 = inside field result region

    loop {
        match reader.read_event()? {
            Event::Start(e) => {
                let local = local_name(e.name().as_ref());
                match local.as_str() {
                    "pPr" if depth == 0 => {
                        let (s, explicit_align, sid, npr, spr) = parse_paragraph_properties(reader)?;
                        style = s;
                        if let Some(a) = explicit_align {
                            alignment = a;
                        }
                        style_id = sid;
                        num_pr_ref = npr;
                        para_sect_pr = spr;
                    }
                    "r" if depth == 0 => {
                        let (mut run, dr) = parse_run(reader, ctx, styles, None)?;
                        // Track field state: fldChar begin/separate/end spans across runs
                        if run.text.contains('\u{FFFE}') {
                            // Marker for fldChar separate (set in parse_run)
                            run.text = run.text.replace('\u{FFFE}', "");
                            field_result_depth += 1;
                        }
                        if run.text.contains('\u{FFFF}') {
                            // Marker for fldChar end
                            run.text = run.text.replace('\u{FFFF}', "");
                            field_result_depth -= 1;
                        }
                        // Suppress cached field result text (between separate and end)
                        // when the field was already evaluated (e.g. PAGE → "#")
                        if field_result_depth > 0 && run.field_type.is_none() {
                            run.text.clear();
                        }
                        runs.push(run);
                        if let Some(drawing) = dr {
                            if let Some(image) = drawing.image {
                                images.push(image);
                            }
                            if let Some(shape) = drawing.shape {
                                found_shapes.push(shape);
                            }
                            if let Some(tb) = drawing.text_box {
                                found_text_boxes.push(tb);
                            }
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
                        let hyperlink_runs = parse_hyperlink_runs(reader, ctx, styles, link_url)?;
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
                        let tracked_runs = parse_tracked_change_runs(reader, ctx, styles, "ins", tc)?;
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
                        let tracked_runs = parse_tracked_change_runs(reader, ctx, styles, "del", tc)?;
                        runs.extend(tracked_runs);
                    }
                    // mc:AlternateContent at paragraph level
                    // OOXML spec (ECMA-376 Part 3): process mc:Choice only, skip mc:Fallback.
                    // mc:Choice may contain both drawings AND text runs (w:r) that belong to this paragraph.
                    "AlternateContent" if depth == 0 => {
                        let mut ac_depth = 0u32;
                        let mut in_choice = false;
                        let mut in_fallback = false;
                        loop {
                            match reader.read_event()? {
                                Event::Start(se) => {
                                    let sl = local_name(se.name().as_ref());
                                    if sl == "Choice" && ac_depth == 0 {
                                        in_choice = true;
                                        ac_depth += 1;
                                    } else if sl == "Fallback" && ac_depth == 0 {
                                        in_fallback = true;
                                        ac_depth += 1;
                                    } else if in_choice && sl == "drawing" && ac_depth == 1 {
                                        let dr = parse_drawing(reader, ctx, styles)?;
                                        if let Some(image) = dr.image {
                                            images.push(image);
                                        }
                                        if let Some(shape) = dr.shape {
                                            found_shapes.push(shape);
                                        }
                                        if let Some(tb) = dr.text_box {
                                            found_text_boxes.push(tb);
                                        }
                                    } else if in_choice && sl == "r" && ac_depth == 1 {
                                        // Text runs inside mc:Choice belong to this paragraph
                                        let (run, dr) = parse_run(reader, ctx, styles, None)?;
                                        runs.push(run);
                                        if let Some(drawing) = dr {
                                            if let Some(image) = drawing.image {
                                                images.push(image);
                                            }
                                            if let Some(shape) = drawing.shape {
                                                found_shapes.push(shape);
                                            }
                                            if let Some(tb) = drawing.text_box {
                                                found_text_boxes.push(tb);
                                            }
                                        }
                                    } else {
                                        ac_depth += 1;
                                    }
                                }
                                Event::End(ee) => {
                                    let sl = local_name(ee.name().as_ref());
                                    if sl == "AlternateContent" && ac_depth == 0 {
                                        break;
                                    }
                                    if sl == "Choice" && in_choice { in_choice = false; }
                                    if sl == "Fallback" && in_fallback { in_fallback = false; }
                                    if ac_depth > 0 { ac_depth -= 1; }
                                }
                                Event::Eof => break,
                                _ => {}
                            }
                        }
                    }
                    // Inline Structured Document Tag (w:sdt inside w:p)
                    // ECMA-376: sdtContent contains runs that are part of the paragraph
                    "sdt" if depth == 0 => {
                        let mut sdt_depth = 1u32;
                        let mut in_sdt_content = false;
                        loop {
                            match reader.read_event()? {
                                Event::Start(se) => {
                                    let sl = local_name(se.name().as_ref());
                                    if sl == "sdtContent" && sdt_depth == 1 {
                                        in_sdt_content = true;
                                    } else if in_sdt_content && sl == "r" {
                                        let (run, dr) = parse_run(reader, ctx, styles, None)?;
                                        runs.push(run);
                                        if let Some(drawing) = dr {
                                            if let Some(image) = drawing.image {
                                                images.push(image);
                                            }
                                            if let Some(shape) = drawing.shape {
                                                found_shapes.push(shape);
                                            }
                                            if let Some(tb) = drawing.text_box {
                                                found_text_boxes.push(tb);
                                            }
                                        }
                                    } else {
                                        sdt_depth += 1;
                                    }
                                }
                                Event::End(ee) => {
                                    let sl = local_name(ee.name().as_ref());
                                    if sl == "sdtContent" {
                                        in_sdt_content = false;
                                    } else if sl == "sdt" && sdt_depth == 1 {
                                        break;
                                    } else if sdt_depth > 1 {
                                        sdt_depth -= 1;
                                    }
                                }
                                Event::Eof => break,
                                _ => {}
                            }
                        }
                    }
                    // OMML math expressions
                    "oMathPara" | "oMath" if depth == 0 => {
                        let math_text = parse_omml(reader, &local)?;
                        if !math_text.is_empty() {
                            runs.push(Run {
                                text: math_text,
                                style: RunStyle { font_family: Some("Cambria Math".to_string()), ..RunStyle::default() },
                                url: None,
                                footnote_ref: None,
                                endnote_ref: None,
                                comment_range_start: Vec::new(),
                                comment_range_end: Vec::new(),
                                tracked_change: None,
                                ruby: None,
                                bookmark_name: None,
                                is_math: true,
                                field_type: None,
                            });
                        }
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
                        "bookmarkStart" => {
                            let mut bk_name = None;
                            for attr in e.attributes().flatten() {
                                let key = local_name(attr.key.as_ref());
                                if key == "name" {
                                    let val = String::from_utf8_lossy(&attr.value).to_string();
                                    if val != "_GoBack" {
                                        bk_name = Some(val);
                                    }
                                }
                            }
                            if let Some(name) = bk_name {
                                // Create an empty anchor run for the bookmark
                                runs.push(Run {
                                    text: String::new(),
                                    style: RunStyle::default(),
                                    url: None,
                                    footnote_ref: None,
                                    endnote_ref: None,
                                    comment_range_start: Vec::new(),
                                    comment_range_end: Vec::new(),
                                    tracked_change: None,
                                    ruby: None,
                                    bookmark_name: Some(name),
                                    is_math: false,
                                    field_type: None,
                                });
                            }
                        }
                        "bookmarkEnd" => {
                            // End marker; anchor is already placed at bookmarkStart
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
    // ECMA-376: paragraph with no pStyle implicitly uses the default paragraph style (w:default="1")
    let effective_style_id = style_id.clone()
        .or_else(|| styles.default_paragraph_style_id.clone());
    if let Some(ref sid) = effective_style_id {
        if let Some(defined) = styles.styles.get(sid) {
            // Inherit alignment from style if not explicitly set in paragraph
            if alignment == Alignment::default() {
                if let Some(style_align) = defined.alignment {
                    alignment = style_align;
                }
            }
            let ds = &defined.paragraph;
            if style.space_before.is_none() {
                style.space_before = ds.space_before;
            }
            if style.space_after.is_none() {
                style.space_after = ds.space_after;
            }
            if style.line_spacing.is_none() {
                style.line_spacing = ds.line_spacing;
                style.line_spacing_rule = ds.line_spacing_rule.clone();
            }
            // Merge style's default_run_style field-by-field (style definition takes priority for unset fields)
            if let Some(ref style_rs) = ds.default_run_style {
                if let Some(ref mut para_rs) = style.default_run_style {
                    // Merge: style definition fills in missing fields
                    if para_rs.font_size.is_none() { para_rs.font_size = style_rs.font_size; }
                    if para_rs.font_family.is_none() { para_rs.font_family = style_rs.font_family.clone(); }
                    if para_rs.font_family_east_asia.is_none() { para_rs.font_family_east_asia = style_rs.font_family_east_asia.clone(); }
                    if !para_rs.has_explicit_east_asia && style_rs.has_explicit_east_asia { para_rs.has_explicit_east_asia = true; }
                    if para_rs.color.is_none() { para_rs.color = style_rs.color.clone(); }
                    if !para_rs.bold { para_rs.bold = style_rs.bold; }
                    if !para_rs.italic { para_rs.italic = style_rs.italic; }
                } else {
                    style.default_run_style = ds.default_run_style.clone();
                }
            }
            // Inherit keepNext, keepLines, contextualSpacing from style
            if ds.keep_next { style.keep_next = true; }
            if ds.keep_lines { style.keep_lines = true; }
            if ds.contextual_spacing { style.contextual_spacing = true; }
            // Inherit numPr from style definition
            if style.num_id.is_none() {
                if let Some(ref nid) = ds.num_id {
                    style.num_id = Some(nid.clone());
                    style.num_ilvl = ds.num_ilvl;
                }
            }
            // Inherit tab stops from style
            if style.tab_stops.is_empty() && !ds.tab_stops.is_empty() {
                style.tab_stops = ds.tab_stops.clone();
            }
            // Inherit paragraph borders from style
            if style.borders.is_none() {
                style.borders = ds.borders.clone();
            }
            // Inherit indents from style
            if style.indent_left.is_none() {
                style.indent_left = ds.indent_left;
            }
            if style.indent_right.is_none() {
                style.indent_right = ds.indent_right;
            }
            if style.indent_first_line.is_none() {
                style.indent_first_line = ds.indent_first_line;
            }
            // Inherit shading from style
            if style.shading.is_none() {
                style.shading = ds.shading.clone();
            }
            // Inherit page_break_before from style
            if ds.page_break_before {
                style.page_break_before = true;
            }
            // Inherit snap_to_grid from style (false overrides struct default true).
            // Round 29: footnote text style "a8" sets snapToGrid=0; without this
            // inheritance, footnote paragraphs were grid-snapped to body line
            // pitch, causing wide line spacing in the footnote area.
            if !ds.snap_to_grid {
                style.snap_to_grid = false;
            }
        }
    }

    // Apply docDefaults fallback (field-by-field merge per ECMA-376 priority)
    if let Some(ref doc_rs) = styles.doc_default_run_style {
        if let Some(ref mut para_rs) = style.default_run_style {
            if para_rs.font_size.is_none() { para_rs.font_size = doc_rs.font_size; }
            if para_rs.font_family.is_none() { para_rs.font_family = doc_rs.font_family.clone(); }
            if para_rs.font_family_east_asia.is_none() { para_rs.font_family_east_asia = doc_rs.font_family_east_asia.clone(); }
            if !para_rs.has_explicit_east_asia && doc_rs.has_explicit_east_asia { para_rs.has_explicit_east_asia = true; }
            if para_rs.color.is_none() { para_rs.color = doc_rs.color.clone(); }
        } else {
            style.default_run_style = styles.doc_default_run_style.clone();
        }
    }
    if let Some(ref doc_para) = styles.doc_default_para_style {
        if style.space_before.is_none() {
            style.space_before = doc_para.space_before;
        }
        if style.space_after.is_none() {
            style.space_after = doc_para.space_after;
        }
        if style.before_lines.is_none() {
            style.before_lines = doc_para.before_lines;
        }
        if style.after_lines.is_none() {
            style.after_lines = doc_para.after_lines;
        }
        if style.line_spacing.is_none() {
            style.line_spacing = doc_para.line_spacing;
            style.line_spacing_rule = doc_para.line_spacing_rule.clone();
            style.line_spacing_from_doc_defaults = true;
        }
        if style.indent_left.is_none() {
            style.indent_left = doc_para.indent_left;
        }
        if style.indent_right.is_none() {
            style.indent_right = doc_para.indent_right;
        }
        if style.indent_first_line.is_none() {
            style.indent_first_line = doc_para.indent_first_line;
        }
        // COM-confirmed: pPrDefault widowControl=false must override the struct default (true)
        if !style.has_explicit_widow_control {
            style.widow_control = doc_para.widow_control;
        }
    }
    // Inherit alignment from docDefaults pPrDefault (jc)
    if alignment == Alignment::default() {
        if let Some(doc_align) = styles.doc_default_alignment {
            alignment = doc_align;
        }
    }

    // Inherit numPr from style definition if paragraph doesn't have its own
    if num_pr_ref.is_none() {
        if let Some(ref nid) = style.num_id {
            if !nid.is_empty() && nid != "0" {
                num_pr_ref = Some(NumPrRef {
                    num_id: nid.clone(),
                    ilvl: style.num_ilvl,
                });
            }
        }
    }

    // Resolve list marker from numbering definitions
    if let Some(npr) = num_pr_ref {
        if !npr.num_id.is_empty() && npr.num_id != "0" {
            let resolved = ctx.numbering.resolve_marker_full(
                &npr.num_id,
                npr.ilvl,
                &mut ctx.list_counters.borrow_mut(),
            );
            style.list_marker = Some(resolved.text);
            style.list_suff = Some(resolved.suff);
            style.list_tab_stop = resolved.tab_stop;
            if let Some(ind) = resolved.hanging {
                style.list_indent = Some(ind);
                if style.indent_left.is_none() {
                    if let Some(left) = ctx.numbering.get_level_indent(&npr.num_id, npr.ilvl) {
                        style.indent_left = Some(left);
                    }
                }
            }
        }
    }

    // Convert inline page break (w:br type="page" as \x0C in first run) to page_break_before
    if let Some(first_run) = runs.first() {
        if first_run.text.trim() == "\x0C" || first_run.text == "\x0C" {
            style.page_break_before = true;
            runs.remove(0); // Remove the break-only run
        }
    }

    // Store style ID for contextual spacing comparison
    style.style_id = style_id;

    // Separate inline images (no position) from floating images
    let mut inline_images: Vec<Block> = Vec::new();
    let mut floating_imgs: Vec<Image> = Vec::new();
    for img in images {
        if img.position.is_some() {
            floating_imgs.push(img);
        } else {
            inline_images.push(Block::Image(img));
        }
    }

    Ok(ParagraphResult {
        paragraph: Paragraph {
            runs,
            style,
            alignment,
            shapes: found_shapes.clone(),
        },
        sect_pr: para_sect_pr,
        shapes: Vec::new(), // shapes are now in Paragraph.shapes, not page-level
        text_boxes: found_text_boxes,
        inline_images,
        floating_images: floating_imgs,
    })
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
) -> Result<(ParagraphStyle, Option<Alignment>, Option<String>, Option<NumPrRef>, Option<SectionProperties>), ParseError> {
    let mut style = ParagraphStyle::default();
    let mut alignment: Option<Alignment> = None;
    let mut style_id: Option<String> = None;
    let mut num_pr: Option<NumPrRef> = None;
    let mut sect_pr: Option<SectionProperties> = None;
    let mut has_explicit_widow_control = false;
    let mut depth = 0;

    loop {
        match reader.read_event()? {
            Event::Start(e) => {
                let local = local_name(e.name().as_ref());
                match local.as_str() {
                    "rPr" if depth == 0 => {
                        // pPr/rPr: paragraph-level default run properties (empty para font)
                        let mut ppr_rpr = RunStyle::default();
                        let mut rd = 0;
                        loop {
                            match reader.read_event()? {
                                Event::Empty(e2) => {
                                    let l = local_name(e2.name().as_ref());
                                    if l == "sz" {
                                        for a in e2.attributes().flatten() {
                                            if local_name(a.key.as_ref()) == "val" {
                                                if let Ok(v) = std::str::from_utf8(&a.value) {
                                                    if let Ok(hp) = v.parse::<f32>() { ppr_rpr.font_size = Some(hp / 2.0); }
                                                }
                                            }
                                        }
                                    } else if l == "rFonts" {
                                        for a in e2.attributes().flatten() {
                                            let k = local_name(a.key.as_ref());
                                            let v = String::from_utf8_lossy(&a.value).to_string();
                                            match k.as_str() {
                                                "ascii" | "hAnsi" => { ppr_rpr.font_family = Some(v); }
                                                "eastAsia" => {
                                                    ppr_rpr.font_family_east_asia = Some(v);
                                                    ppr_rpr.has_explicit_east_asia = true;
                                                }
                                                _ => {}
                                            }
                                        }
                                    } else if l == "b" { ppr_rpr.bold = true; }
                                }
                                Event::Start(_) => { rd += 1; }
                                Event::End(_) => { if rd == 0 { break; } rd -= 1; }
                                Event::Eof => break,
                                _ => {}
                            }
                        }
                        style.ppr_rpr = Some(ppr_rpr);
                    }
                    "numPr" if depth == 0 => {
                        num_pr = Some(parse_num_pr(reader)?);
                    }
                    "spacing" if depth == 0 => {
                        style.has_direct_spacing = true;
                        let mut line_val: Option<f32> = None;
                        let mut line_rule: Option<String> = None;
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
                                "beforeLines" => {
                                    style.before_lines = val.parse::<f32>().ok();
                                }
                                "afterLines" => {
                                    style.after_lines = val.parse::<f32>().ok();
                                }
                                "line" => {
                                    line_val = val.parse::<f32>().ok();
                                }
                                "lineRule" => {
                                    line_rule = Some(val.to_string());
                                }
                                _ => {}
                            }
                        }
                        if let Some(lv) = line_val {
                            match line_rule.as_deref() {
                                Some("exact") => {
                                    style.line_spacing = Some(lv / 20.0);
                                    style.line_spacing_rule = Some("exact".to_string());
                                }
                                Some("atLeast") => {
                                    style.line_spacing = Some(lv / 20.0);
                                    style.line_spacing_rule = Some("atLeast".to_string());
                                }
                                _ => {
                                    style.line_spacing = Some(lv / 240.0);
                                }
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
                                alignment = Some(match val.as_ref() {
                                    "left" | "start" => Alignment::Left,
                                    "center" => Alignment::Center,
                                    "right" | "end" => Alignment::Right,
                                    "both" => Alignment::Justify,
                                    "distribute" => Alignment::Distribute,
                                    _ => Alignment::Left,
                                });
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
                        let mut line_val: Option<f32> = None;
                        let mut line_rule: Option<String> = None;
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
                                "beforeLines" => {
                                    style.before_lines = val.parse::<f32>().ok();
                                }
                                "afterLines" => {
                                    style.after_lines = val.parse::<f32>().ok();
                                }
                                "line" => {
                                    line_val = val.parse::<f32>().ok();
                                }
                                "lineRule" => {
                                    line_rule = Some(val.to_string());
                                }
                                _ => {}
                            }
                        }
                        if let Some(lv) = line_val {
                            match line_rule.as_deref() {
                                Some("exact") => {
                                    // Exact: value in twips, convert to points
                                    style.line_spacing = Some(lv / 20.0);
                                    style.line_spacing_rule = Some("exact".to_string());
                                }
                                Some("atLeast") => {
                                    // At least: value in twips, convert to points
                                    style.line_spacing = Some(lv / 20.0);
                                    style.line_spacing_rule = Some("atLeast".to_string());
                                }
                                _ => {
                                    // Auto: proportional, divide by 240
                                    style.line_spacing = Some(lv / 240.0);
                                }
                            }
                        }
                    }
                    "ind" => {
                        // Track leftChars/rightChars for override logic
                        // Must be outside attr loop: leftChars overrides left regardless of XML attr order
                        let mut left_chars: Option<f32> = None;
                        let mut right_chars: Option<f32> = None;
                        let mut first_line_chars: Option<f32> = None;
                        for attr in e.attributes().flatten() {
                            let key = local_name(attr.key.as_ref());
                            let val = String::from_utf8_lossy(&attr.value);
                            match key.as_str() {
                                "left" | "start" => {
                                    style.indent_left =
                                        val.parse::<f32>().ok().map(|v| v / 20.0);
                                }
                                "right" | "end" => {
                                    style.indent_right =
                                        val.parse::<f32>().ok().map(|v| v / 20.0);
                                }
                                "leftChars" | "startChars" => {
                                    // COM-confirmed: leftChars overrides left (not additive)
                                    // charWidth = 10.5pt (Japanese default)
                                    left_chars = val.parse::<f32>().ok();
                                }
                                "rightChars" | "endChars" => {
                                    right_chars = val.parse::<f32>().ok();
                                }
                                "firstLine" => {
                                    style.indent_first_line =
                                        val.parse::<f32>().ok().map(|v| v / 20.0);
                                }
                                "firstLineChars" => {
                                    first_line_chars = val.parse::<f32>().ok();
                                }
                                "hanging" => {
                                    // Hanging indent: negative first-line indent
                                    style.indent_first_line =
                                        val.parse::<f32>().ok().map(|v| -(v / 20.0));
                                }
                                _ => {}
                            }
                        }
                        // *Chars attributes override twip values regardless of XML attr order
                        if let Some(lc) = left_chars {
                            style.indent_left = Some(lc / 100.0 * 10.5);
                        }
                        if let Some(rc) = right_chars {
                            style.indent_right = Some(rc / 100.0 * 10.5);
                        }
                        if let Some(fc) = first_line_chars {
                            style.indent_first_line = Some(fc / 100.0 * 10.5);
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
                        has_explicit_widow_control = true;
                    }
                    "bidi" => {
                        let mut enabled = true;
                        for attr in e.attributes().flatten() {
                            if local_name(attr.key.as_ref()) == "val" {
                                let val = String::from_utf8_lossy(&attr.value);
                                enabled = val.as_ref() != "0" && val.as_ref() != "false";
                            }
                        }
                        style.bidi = enabled;
                    }
                    "autoSpaceDE" => {
                        let mut enabled = true;
                        for attr in e.attributes().flatten() {
                            if local_name(attr.key.as_ref()) == "val" {
                                let val = String::from_utf8_lossy(&attr.value);
                                enabled = val.as_ref() != "0" && val.as_ref() != "false";
                            }
                        }
                        style.auto_space_de = enabled;
                    }
                    "autoSpaceDN" => {
                        let mut enabled = true;
                        for attr in e.attributes().flatten() {
                            if local_name(attr.key.as_ref()) == "val" {
                                let val = String::from_utf8_lossy(&attr.value);
                                enabled = val.as_ref() != "0" && val.as_ref() != "false";
                            }
                        }
                        style.auto_space_dn = enabled;
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

    style.has_explicit_widow_control = has_explicit_widow_control;
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
pub(crate) fn parse_paragraph_borders(reader: &mut Reader<&[u8]>) -> Result<ParagraphBorders, ParseError> {
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
    let mut space: f32 = 0.0;

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
                if val == "auto" {
                    color = Some("000000".to_string());
                } else {
                    color = Some(val);
                }
            }
            "space" => {
                space = val.parse::<f32>().unwrap_or(0.0);
            }
            _ => {}
        }
    }

    if style.is_empty() {
        return None;
    }

    Some(BorderDef { style, width, color, space })
}

/// Parse w:tabs element containing w:tab children
pub(crate) fn parse_tab_stops(reader: &mut Reader<&[u8]>) -> Result<Vec<TabStop>, ParseError> {
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
fn parse_run(reader: &mut Reader<&[u8]>, ctx: &ParseContext, styles: &StyleSheet, url: Option<String>) -> Result<(Run, Option<DrawingResult>), ParseError> {
    let mut text = String::new();
    let mut style = RunStyle::default();
    let mut drawing_result: Option<DrawingResult> = None;
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
                        style = parse_run_properties(reader, ctx, styles)?;
                    }
                    "t" | "delText" if depth == 0 => {
                        in_text = true;
                    }
                    "instrText" if depth == 0 => {
                        in_instr_text = true;
                    }
                    "drawing" if depth == 0 => {
                        drawing_result = Some(parse_drawing(reader, ctx, styles)?);
                    }
                    // VML legacy picture/shape
                    "pict" if depth == 0 => {
                        let vml = parse_vml_pict(reader, ctx, styles)?;
                        if drawing_result.is_none() {
                            drawing_result = Some(vml);
                        }
                    }
                    // mc:AlternateContent — prefer Choice (DrawingML)
                    "AlternateContent" if depth == 0 => {
                        let ac = parse_alternate_content(reader, ctx, styles)?;
                        if drawing_result.is_none() {
                            drawing_result = ac;
                        }
                    }
                    // OLE object — extract preview image from embedded VML shape
                    "object" if depth == 0 => {
                        let ole = parse_ole_object(reader, ctx)?;
                        if drawing_result.is_none() {
                            drawing_result = Some(ole);
                        }
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
                if local == "t" || local == "delText" {
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
                        // Track field state via marker characters
                        // (parsed by parent parse_paragraph to manage field_result_depth)
                        for attr in e.attributes().flatten() {
                            if local_name(attr.key.as_ref()) == "fldCharType" {
                                let val = String::from_utf8_lossy(&attr.value);
                                match val.as_ref() {
                                    "separate" => text.push('\u{FFFE}'),
                                    "end" => text.push('\u{FFFF}'),
                                    _ => {} // "begin" — no marker needed
                                }
                            }
                        }
                    }
                    "footnoteReference" => {
                        for attr in e.attributes().flatten() {
                            let key = local_name(attr.key.as_ref());
                            if key == "id" {
                                let val = String::from_utf8_lossy(&attr.value);
                                if let Ok(id) = val.parse::<u32>() {
                                    if id > 0 { // Skip separator/continuation notes (id=0)
                                        footnote_ref = Some(id);
                                        // Word renders just the number (e.g. "1"),
                                        // not "[1]". renumber_note_refs in parse_body
                                        // will rewrite this to the section-local seq.
                                        text = format!("{}", id);
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
    let mut field_type: Option<FieldType> = None;
    if !instr_text.is_empty() {
        let field = instr_text.trim();
        if field.contains("PAGE") && !field.contains("NUMPAGES") {
            text = "#".to_string();
            field_type = Some(FieldType::Page);
        } else if field.contains("NUMPAGES") || field.contains("SECTIONPAGES") {
            text = "#".to_string();
            field_type = Some(FieldType::NumPages);
        } else if field.contains("DATE") || field.contains("TIME") {
            text = field.to_string();
        } else if field.contains("TOC") || field.contains("HYPERLINK") {
            // Table of contents / hyperlink fields — keep existing text (result display)
        } else if field.contains("REF") || field.contains("NOTEREF") || field.contains("PAGEREF") {
            // Cross-reference fields — show placeholder
            if text.is_empty() {
                text = "#".to_string();
            }
        } else if field.contains("AUTHOR") || field.contains("TITLE") || field.contains("SUBJECT") {
            // Document property fields — show field name as placeholder
            if text.is_empty() {
                text = format!("[{}]", field.split_whitespace().next().unwrap_or(field));
            }
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
        bookmark_name: None,
        is_math: false,
        field_type,
    }, drawing_result))
}

/// Parse runs inside a w:hyperlink element
fn parse_hyperlink_runs(reader: &mut Reader<&[u8]>, ctx: &ParseContext, styles: &StyleSheet, url: Option<String>) -> Result<Vec<Run>, ParseError> {
    let mut runs = Vec::new();
    let mut depth = 0;

    loop {
        match reader.read_event()? {
            Event::Start(e) => {
                let local = local_name(e.name().as_ref());
                if local == "r" && depth == 0 {
                    let (run, _dr) = parse_run(reader, ctx, styles, url.clone())?;
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
fn parse_notes_xml(xml: &str, styles: &StyleSheet) -> Result<HashMap<String, Vec<Block>>, ParseError> {
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
                        let pr = parse_paragraph(&mut reader, &note_ctx, styles)?;
                        let para = pr.paragraph;
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

/// Result from parsing a w:drawing element — may contain image, shape, and/or text box
struct DrawingResult {
    image: Option<Image>,
    shape: Option<Shape>,
    text_box: Option<TextBox>,
}

impl DrawingResult {
    /// Returns true if at least one component (image, shape, or text_box) is present
    fn has_content(&self) -> bool {
        self.image.is_some() || self.shape.is_some() || self.text_box.is_some()
    }
}

/// Parse DrawingML color modifier child elements (lumMod, lumOff, tint, shade, etc.)
/// and apply them to a base hex color. Consumes elements until the closing tag.
fn parse_color_modifiers(reader: &mut Reader<&[u8]>, base_hex: &str, end_tag: &str) -> String {
    let mut lum_mod: Option<f32> = None;
    let mut lum_off: Option<f32> = None;
    let mut tint: Option<f32> = None;
    let mut shade: Option<f32> = None;
    let mut sat_mod: Option<f32> = None;

    loop {
        match reader.read_event() {
            Ok(Event::Empty(e)) => {
                let local = local_name(e.name().as_ref());
                for attr in e.attributes().flatten() {
                    if local_name(attr.key.as_ref()) == "val" {
                        let val = String::from_utf8_lossy(&attr.value);
                        let v = val.parse::<f32>().unwrap_or(0.0) / 100000.0;
                        match local.as_str() {
                            "lumMod" => lum_mod = Some(v),
                            "lumOff" => lum_off = Some(v),
                            "tint" => tint = Some(v),
                            "shade" => shade = Some(v),
                            "satMod" => sat_mod = Some(v),
                            "alpha" => {} // Recognized but not yet applied (opacity)
                            _ => {}
                        }
                    }
                }
            }
            Ok(Event::End(e)) => {
                if local_name(e.name().as_ref()) == end_tag {
                    break;
                }
            }
            Ok(Event::Eof) => break,
            _ => {}
        }
    }

    // Apply modifiers to base color
    let r = u8::from_str_radix(&base_hex[0..2], 16).unwrap_or(0) as f32 / 255.0;
    let g = u8::from_str_radix(&base_hex[2..4], 16).unwrap_or(0) as f32 / 255.0;
    let b = u8::from_str_radix(&base_hex[4..6], 16).unwrap_or(0) as f32 / 255.0;

    let (mut r, mut g, mut b) = (r, g, b);

    // Apply tint (move towards white)
    if let Some(t) = tint {
        r = r * t + (1.0 - t);
        g = g * t + (1.0 - t);
        b = b * t + (1.0 - t);
    }

    // Apply shade (move towards black)
    if let Some(s) = shade {
        r *= s;
        g *= s;
        b *= s;
    }

    // Apply HLS modifiers: lumMod/lumOff on L, satMod on S (matches Word output)
    if lum_mod.is_some() || lum_off.is_some() || sat_mod.is_some() {
        // RGB -> HSL
        let max_c = r.max(g).max(b);
        let min_c = r.min(g).min(b);
        let d = max_c - min_c;
        let l_orig = (max_c + min_c) / 2.0;
        let mut l = l_orig;
        // Apply lumMod/lumOff to L
        if lum_mod.is_some() || lum_off.is_some() {
            let m = lum_mod.unwrap_or(1.0);
            let o = lum_off.unwrap_or(0.0);
            l = (l * m + o).clamp(0.0, 1.0);
        }
        if d.abs() < 1e-6 {
            // Achromatic — satMod has no effect
            r = l; g = l; b = l;
        } else {
            let s_orig = if l_orig > 0.5 { d / (2.0 - max_c - min_c).max(1e-6) } else { d / (max_c + min_c).max(1e-6) };
            let h = if (max_c - r).abs() < 1e-6 {
                (g - b) / d + if g < b { 6.0 } else { 0.0 }
            } else if (max_c - g).abs() < 1e-6 {
                (b - r) / d + 2.0
            } else {
                (r - g) / d + 4.0
            } / 6.0;
            // Apply satMod to S
            let s = if let Some(sm) = sat_mod {
                (s_orig * sm).clamp(0.0, 1.0)
            } else {
                s_orig
            };
            // HSL -> RGB with modified L and S, original H
            let hue_to_rgb = |p: f32, q: f32, mut t: f32| -> f32 {
                if t < 0.0 { t += 1.0; }
                if t > 1.0 { t -= 1.0; }
                if t < 1.0/6.0 { return p + (q - p) * 6.0 * t; }
                if t < 1.0/2.0 { return q; }
                if t < 2.0/3.0 { return p + (q - p) * (2.0/3.0 - t) * 6.0; }
                p
            };
            let q = if l < 0.5 { l * (1.0 + s) } else { l + s - l * s };
            let p = 2.0 * l - q;
            r = hue_to_rgb(p, q, h + 1.0/3.0);
            g = hue_to_rgb(p, q, h);
            b = hue_to_rgb(p, q, h - 1.0/3.0);
        }
    }

    format!("{:02X}{:02X}{:02X}",
        (r * 255.0).round() as u8,
        (g * 255.0).round() as u8,
        (b * 255.0).round() as u8,
    )
}

/// Parse a w:drawing element to extract image, shape, or text box info
fn parse_drawing(reader: &mut Reader<&[u8]>, ctx: &ParseContext, styles: &StyleSheet) -> Result<DrawingResult, ParseError> {
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
    let mut h_align: Option<String> = None;
    let mut v_align: Option<String> = None;
    let mut in_align = false;
    let mut wrap_type: Option<WrapType> = None;
    let mut crop: Option<ImageCrop> = None;
    let mut in_pos_h = false;
    let mut in_pos_v = false;
    let mut in_pos_offset = false;
    // Shape properties
    let mut shape_type: Option<String> = None;
    let mut shape_fill: Option<String> = None;
    let mut stroke_color: Option<String> = None;
    let mut stroke_width: Option<f32> = None;
    let mut shape_text_blocks: Vec<Block> = Vec::new();
    let mut rotation: Option<f32> = None;
    let mut has_no_fill = false;
    let mut has_no_stroke = false;
    let mut corner_radius_adj: Option<f32> = None; // avLst adj value (0-100000 scale)
    let mut gradient_stops: Vec<GradientStop> = Vec::new();
    let mut gradient_angle: Option<f32> = None;
    // Text body insets (from bodyPr lIns/rIns/tIns/bIns, in EMU)
    let mut text_inset_left: Option<f32> = None;
    let mut text_inset_right: Option<f32> = None;
    let mut text_inset_top: Option<f32> = None;
    let mut text_inset_bottom: Option<f32> = None;
    let mut text_body_anchor: Option<String> = None;

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
                    // Shape line as Start element — may contain srgbClr child
                    "ln" => {
                        for attr in e.attributes().flatten() {
                            let key = local_name(attr.key.as_ref());
                            if key == "w" {
                                let val = String::from_utf8_lossy(&attr.value);
                                stroke_width = val.parse::<f32>().ok().map(|v| v / 12700.0);
                            }
                        }
                        // Parse children for stroke color
                        let mut ln_depth = 1;
                        loop {
                            match reader.read_event() {
                                Ok(Event::Start(se)) => {
                                    let sl = local_name(se.name().as_ref());
                                    match sl.as_str() {
                                        "schemeClr" => {
                                            let mut base_hex: Option<String> = None;
                                            for attr in se.attributes().flatten() {
                                                if local_name(attr.key.as_ref()) == "val" {
                                                    let val = String::from_utf8_lossy(&attr.value).to_string();
                                                    base_hex = ctx.theme.resolve(&val).cloned();
                                                }
                                            }
                                            if let Some(hex) = base_hex {
                                                stroke_color = Some(parse_color_modifiers(reader, &hex, "schemeClr"));
                                            }
                                            // parse_color_modifiers consumed End tag, no depth change
                                        }
                                        "srgbClr" => {
                                            let mut hex = String::new();
                                            for attr in se.attributes().flatten() {
                                                if local_name(attr.key.as_ref()) == "val" {
                                                    hex = String::from_utf8_lossy(&attr.value).to_string();
                                                }
                                            }
                                            if !hex.is_empty() {
                                                stroke_color = Some(parse_color_modifiers(reader, &hex, "srgbClr"));
                                            }
                                        }
                                        "sysClr" => {
                                            let mut hex = String::new();
                                            for attr in se.attributes().flatten() {
                                                if local_name(attr.key.as_ref()) == "lastClr" {
                                                    hex = String::from_utf8_lossy(&attr.value).to_string();
                                                }
                                            }
                                            if !hex.is_empty() {
                                                stroke_color = Some(parse_color_modifiers(reader, &hex, "sysClr"));
                                            }
                                        }
                                        _ => ln_depth += 1,
                                    }
                                }
                                Ok(Event::Empty(se)) => {
                                    let sl = local_name(se.name().as_ref());
                                    if sl == "srgbClr" {
                                        for attr in se.attributes().flatten() {
                                            if local_name(attr.key.as_ref()) == "val" {
                                                stroke_color = Some(String::from_utf8_lossy(&attr.value).to_string());
                                            }
                                        }
                                    } else if sl == "sysClr" {
                                        for attr in se.attributes().flatten() {
                                            if local_name(attr.key.as_ref()) == "lastClr" {
                                                stroke_color = Some(String::from_utf8_lossy(&attr.value).to_string());
                                            }
                                        }
                                    } else if sl == "schemeClr" {
                                        for attr in se.attributes().flatten() {
                                            if local_name(attr.key.as_ref()) == "val" {
                                                let val = String::from_utf8_lossy(&attr.value).to_string();
                                                if let Some(resolved) = ctx.theme.resolve(&val) {
                                                    stroke_color = Some(resolved.clone());
                                                }
                                            }
                                        }
                                    } else if sl == "noFill" {
                                        has_no_stroke = true;
                                    }
                                }
                                Ok(Event::End(se)) => {
                                    ln_depth -= 1;
                                    if ln_depth == 0 {
                                        break;
                                    }
                                }
                                Ok(Event::Eof) => break,
                                _ => {}
                            }
                        }
                        // We consumed ln's End, decrement outer depth
                        depth -= 1;
                    }
                    // Gradient fill
                    "gradFill" => {
                        let mut gf_depth = 1;
                        let mut current_gs_pos: Option<f32> = None;
                        loop {
                            match reader.read_event() {
                                Ok(Event::Start(se)) => {
                                    let sl = local_name(se.name().as_ref());
                                    gf_depth += 1;
                                    if sl == "gs" {
                                        for attr in se.attributes().flatten() {
                                            if local_name(attr.key.as_ref()) == "pos" {
                                                let val = String::from_utf8_lossy(&attr.value);
                                                current_gs_pos = val.parse::<f32>().ok().map(|v| v / 1000.0);
                                            }
                                        }
                                    } else if sl == "lin" {
                                        for attr in se.attributes().flatten() {
                                            if local_name(attr.key.as_ref()) == "ang" {
                                                let val = String::from_utf8_lossy(&attr.value);
                                                gradient_angle = val.parse::<f32>().ok().map(|v| v / 60000.0);
                                            }
                                        }
                                    }
                                }
                                Ok(Event::Empty(se)) => {
                                    let sl = local_name(se.name().as_ref());
                                    if sl == "srgbClr" {
                                        if let Some(pos) = current_gs_pos {
                                            for attr in se.attributes().flatten() {
                                                if local_name(attr.key.as_ref()) == "val" {
                                                    gradient_stops.push(GradientStop {
                                                        position: pos,
                                                        color: String::from_utf8_lossy(&attr.value).to_string(),
                                                    });
                                                }
                                            }
                                        }
                                    } else if sl == "lin" {
                                        for attr in se.attributes().flatten() {
                                            if local_name(attr.key.as_ref()) == "ang" {
                                                let val = String::from_utf8_lossy(&attr.value);
                                                gradient_angle = val.parse::<f32>().ok().map(|v| v / 60000.0);
                                            }
                                        }
                                    } else if sl == "gs" {
                                        // Empty gs element — unlikely but handle gracefully
                                    }
                                }
                                Ok(Event::End(se)) => {
                                    let sl = local_name(se.name().as_ref());
                                    if sl == "gs" {
                                        current_gs_pos = None;
                                    }
                                    gf_depth -= 1;
                                    if gf_depth == 0 { break; }
                                }
                                Ok(Event::Eof) => break,
                                _ => {}
                            }
                        }
                        depth -= 1; // consumed gradFill End
                    }
                    // schemeClr with child modifiers (lumMod, lumOff, tint, shade)
                    "schemeClr" => {
                        let mut base_hex: Option<String> = None;
                        for attr in e.attributes().flatten() {
                            if local_name(attr.key.as_ref()) == "val" {
                                let val = String::from_utf8_lossy(&attr.value).to_string();
                                base_hex = ctx.theme.resolve(&val).cloned();
                            }
                        }
                        if let Some(hex) = base_hex {
                            let final_color = parse_color_modifiers(reader, &hex, "schemeClr");
                            if shape_fill.is_none() && !has_no_fill {
                                shape_fill = Some(final_color);
                            }
                        }
                        depth -= 1; // consumed schemeClr End
                    }
                    // srgbClr with child modifiers (alpha, tint, shade, lumMod, etc.)
                    "srgbClr" => {
                        let mut hex = String::new();
                        for attr in e.attributes().flatten() {
                            if local_name(attr.key.as_ref()) == "val" {
                                hex = String::from_utf8_lossy(&attr.value).to_string();
                            }
                        }
                        if !hex.is_empty() {
                            let final_color = parse_color_modifiers(reader, &hex, "srgbClr");
                            if shape_fill.is_none() && !has_no_fill {
                                shape_fill = Some(final_color);
                            }
                        }
                        depth -= 1;
                    }
                    // sysClr with child modifiers
                    "sysClr" => {
                        let mut hex = String::new();
                        for attr in e.attributes().flatten() {
                            if local_name(attr.key.as_ref()) == "lastClr" {
                                hex = String::from_utf8_lossy(&attr.value).to_string();
                            }
                        }
                        if !hex.is_empty() {
                            let final_color = parse_color_modifiers(reader, &hex, "sysClr");
                            if shape_fill.is_none() && !has_no_fill {
                                shape_fill = Some(final_color);
                            }
                        }
                        depth -= 1;
                    }
                    // Shape transform rotation (xfrm rot attribute in 60000ths of a degree)
                    "xfrm" => {
                        for attr in e.attributes().flatten() {
                            let key = local_name(attr.key.as_ref());
                            if key == "rot" {
                                let val = String::from_utf8_lossy(&attr.value);
                                rotation = val.parse::<f32>().ok().map(|v| v / 60000.0);
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
                    "align" => {
                        in_align = true;
                    }
                    // Text body properties — text insets (bodyPr as start element with children)
                    "bodyPr" => {
                        for attr in e.attributes().flatten() {
                            let key = local_name(attr.key.as_ref());
                            let val = String::from_utf8_lossy(&attr.value);
                            match key.as_str() {
                                "lIns" => { text_inset_left = val.parse::<f32>().ok().map(|v| v / 12700.0); }
                                "rIns" => { text_inset_right = val.parse::<f32>().ok().map(|v| v / 12700.0); }
                                "tIns" => { text_inset_top = val.parse::<f32>().ok().map(|v| v / 12700.0); }
                                "bIns" => { text_inset_bottom = val.parse::<f32>().ok().map(|v| v / 12700.0); }
                                "anchor" => { text_body_anchor = Some(val.to_string()); }
                                _ => {}
                            }
                        }
                    }
                    // DrawingML shape text content
                    "txbxContent" => {
                        // Parse text blocks inside shape
                        loop {
                            match reader.read_event() {
                                Ok(Event::Start(se)) => {
                                    let sl = local_name(se.name().as_ref());
                                    if sl == "p" {
                                        if let Ok(pr) = parse_paragraph(reader, ctx, styles) {
                                            shape_text_blocks.push(Block::Paragraph(pr.paragraph));
                                        }
                                    }
                                }
                                Ok(Event::End(se)) => {
                                    if local_name(se.name().as_ref()) == "txbxContent" {
                                        break;
                                    }
                                }
                                Ok(Event::Eof) => break,
                                _ => {}
                            }
                        }
                        // We consumed the txbxContent end tag, so decrement depth
                        depth -= 1;
                    }
                    // a:blip as Start element (has child elements like a:extLst)
                    "blip" => {
                        for attr in e.attributes().flatten() {
                            let key = std::str::from_utf8(attr.key.as_ref()).unwrap_or("");
                            if key == "r:embed" || key.ends_with(":embed") || key == "embed" {
                                rel_id = Some(String::from_utf8_lossy(&attr.value).to_string());
                            }
                        }
                    }
                    // Shape preset geometry as Start element (has avLst child)
                    "prstGeom" => {
                        for attr in e.attributes().flatten() {
                            let key = local_name(attr.key.as_ref());
                            if key == "prst" {
                                shape_type = Some(String::from_utf8_lossy(&attr.value).to_string());
                            }
                        }
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
                } else if in_align {
                    let content = e.unescape().unwrap_or_default().to_string();
                    if in_pos_h { h_align = Some(content); }
                    else if in_pos_v { v_align = Some(content); }
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
                    "srcRect" => {
                        // a:srcRect — image crop percentages (in 1/1000th percent)
                        let mut c = ImageCrop { top: 0.0, right: 0.0, bottom: 0.0, left: 0.0 };
                        for attr in e.attributes().flatten() {
                            let key = std::str::from_utf8(attr.key.as_ref()).unwrap_or("");
                            let val = String::from_utf8_lossy(&attr.value);
                            if let Ok(v) = val.parse::<f32>() {
                                let pct = v / 1000.0; // Convert from 1/1000th percent to percent
                                match key {
                                    "t" => c.top = pct,
                                    "r" => c.right = pct,
                                    "b" => c.bottom = pct,
                                    "l" => c.left = pct,
                                    _ => {}
                                }
                            }
                        }
                        if c.top > 0.0 || c.right > 0.0 || c.bottom > 0.0 || c.left > 0.0 {
                            crop = Some(c);
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
                    // Shape preset geometry (e.g. rect, ellipse, roundRect, triangle, etc.)
                    "prstGeom" => {
                        for attr in e.attributes().flatten() {
                            let key = local_name(attr.key.as_ref());
                            if key == "prst" {
                                shape_type = Some(String::from_utf8_lossy(&attr.value).to_string());
                            }
                        }
                    }
                    // Adjustment values inside prstGeom (e.g. roundRect corner radius)
                    "gd" => {
                        let mut is_adj = false;
                        let mut adj_val: Option<f32> = None;
                        for attr in e.attributes().flatten() {
                            let key = local_name(attr.key.as_ref());
                            if key == "name" && String::from_utf8_lossy(&attr.value) == "adj" {
                                is_adj = true;
                            }
                            if key == "fmla" {
                                let fmla = String::from_utf8_lossy(&attr.value).to_string();
                                if let Some(val_str) = fmla.strip_prefix("val ") {
                                    adj_val = val_str.trim().parse::<f32>().ok();
                                }
                            }
                        }
                        if is_adj {
                            if let Some(val) = adj_val {
                                corner_radius_adj = Some(val);
                            }
                        }
                    }
                    // Shape solid fill color
                    "srgbClr" => {
                        for attr in e.attributes().flatten() {
                            let key = local_name(attr.key.as_ref());
                            if key == "val" {
                                let val = String::from_utf8_lossy(&attr.value).to_string();
                                if shape_fill.is_none() && !has_no_fill {
                                    shape_fill = Some(val);
                                }
                            }
                        }
                    }
                    // System color — use lastClr attribute as fallback
                    "sysClr" => {
                        for attr in e.attributes().flatten() {
                            let key = local_name(attr.key.as_ref());
                            if key == "lastClr" {
                                let val = String::from_utf8_lossy(&attr.value).to_string();
                                if shape_fill.is_none() && !has_no_fill {
                                    shape_fill = Some(val);
                                }
                            }
                        }
                    }
                    // Theme color reference — resolve via theme color map
                    "schemeClr" => {
                        for attr in e.attributes().flatten() {
                            let key = local_name(attr.key.as_ref());
                            if key == "val" {
                                let val = String::from_utf8_lossy(&attr.value).to_string();
                                if shape_fill.is_none() && !has_no_fill {
                                    if let Some(resolved) = ctx.theme.resolve(&val) {
                                        shape_fill = Some(resolved.clone());
                                    }
                                }
                            }
                        }
                    }
                    // Text body properties — text insets (bodyPr as empty element)
                    "bodyPr" => {
                        for attr in e.attributes().flatten() {
                            let key = local_name(attr.key.as_ref());
                            let val = String::from_utf8_lossy(&attr.value);
                            match key.as_str() {
                                "lIns" => { text_inset_left = val.parse::<f32>().ok().map(|v| v / 12700.0); }
                                "rIns" => { text_inset_right = val.parse::<f32>().ok().map(|v| v / 12700.0); }
                                "tIns" => { text_inset_top = val.parse::<f32>().ok().map(|v| v / 12700.0); }
                                "bIns" => { text_inset_bottom = val.parse::<f32>().ok().map(|v| v / 12700.0); }
                                "anchor" => { text_body_anchor = Some(val.to_string()); }
                                _ => {}
                            }
                        }
                    }
                    "noFill" => { has_no_fill = true; }
                    "noLn" => { has_no_stroke = true; }
                    // Shape line/outline properties (as empty element)
                    "ln" => {
                        for attr in e.attributes().flatten() {
                            let key = local_name(attr.key.as_ref());
                            if key == "w" {
                                let val = String::from_utf8_lossy(&attr.value);
                                stroke_width = val.parse::<f32>().ok().map(|v| v / 12700.0);
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
                    "align" => { in_align = false; }
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

    let position = if is_anchor {
        Some(FloatingPosition { x: pos_x, y: pos_y, h_relative, v_relative, h_align, v_align })
    } else {
        None
    };

    // Build image if we have a blip reference
    let image = if let Some(rid) = rel_id {
        let data = ctx.media.get(&rid).cloned().unwrap_or_default();
        let content_type = ctx.media_types.get(&rid).cloned();
        Some(Image {
            data,
            width,
            height,
            alt_text,
            content_type,
            position: position.clone(),
            wrap_type,
            crop,
            anchor_block_index: 0,
        })
    } else {
        None
    };

    // Save stroke info for TextBox before Shape takes ownership
    let stroke_color_saved = stroke_color.clone();
    let stroke_width_saved = stroke_width;

    // Build shape if we detected a preset geometry
    let shape = if let Some(ref st) = shape_type {
        Some(Shape {
            shape_type: st.clone(),
            width,
            height,
            position: position.clone(),
            fill: if has_no_fill { None } else { shape_fill.clone() },
            stroke_color: if has_no_stroke { None } else { stroke_color },
            stroke_width: if has_no_stroke { None } else { stroke_width },
            text_blocks: Vec::new(), // text goes to text_box
            rotation,
            gradient_stops: gradient_stops.clone(),
            gradient_angle,
            anchor_block_index: 0,
            v_text_anchor: None,
        })
    } else {
        None
    };

    // Compute corner radius for roundRect shapes
    let corner_radius = if shape_type.as_deref() == Some("roundRect") {
        // adj value in 0-100000 scale; default roundRect adj = 16667 (1/6)
        let adj = corner_radius_adj.unwrap_or(16667.0);
        let min_dim = width.min(height);
        Some(min_dim * adj / 100000.0)
    } else {
        None
    };

    // Build text box if we have text content in a shape
    let text_box = if !shape_text_blocks.is_empty() {
        Some(TextBox {
            blocks: shape_text_blocks,
            width,
            height,
            position,
            border: !has_no_stroke,
            stroke_color: if has_no_stroke { None } else { stroke_color_saved.clone() },
            stroke_width: if has_no_stroke { None } else { stroke_width_saved },
            fill: if has_no_fill { None } else { shape_fill.clone().or_else(|| shape_type.as_ref().map(|_| "FFFFFF".to_string())) },
            anchor_block_index: 0, // set by caller in parse_body
            corner_radius,
            inset_left: text_inset_left,
            inset_right: text_inset_right,
            inset_top: text_inset_top,
            inset_bottom: text_inset_bottom,
            wrap_type,
            v_text_anchor: text_body_anchor,
        })
    } else {
        None
    };

    Ok(DrawingResult { image, shape, text_box })
}

/// Parse VML w:pict element (legacy shapes/images)
fn parse_vml_pict(reader: &mut Reader<&[u8]>, ctx: &ParseContext, styles: &StyleSheet) -> Result<DrawingResult, ParseError> {
    let mut shape_type = None;
    let mut width: f32 = 0.0;
    let mut height: f32 = 0.0;
    let mut fill_color: Option<String> = None;
    let mut stroke_color_val: Option<String> = None;
    let mut stroke_width_val: Option<f32> = None;
    let mut no_stroke = false;
    let mut rel_id: Option<String> = None;
    let mut text_blocks: Vec<Block> = Vec::new();
    let mut v_text_anchor: Option<String> = None;
    let mut depth = 0;

    loop {
        match reader.read_event()? {
            Event::Start(e) => {
                let local = local_name(e.name().as_ref());
                depth += 1;
                match local.as_str() {
                    // VML text box content — consume all paragraphs inside
                    "txbxContent" => {
                        loop {
                            match reader.read_event() {
                                Ok(Event::Start(se)) => {
                                    if local_name(se.name().as_ref()) == "p" {
                                        if let Ok(pr) = parse_paragraph(reader, ctx, styles) {
                                            text_blocks.push(Block::Paragraph(pr.paragraph));
                                        }
                                    }
                                }
                                Ok(Event::End(se)) => {
                                    if local_name(se.name().as_ref()) == "txbxContent" {
                                        break;
                                    }
                                }
                                Ok(Event::Eof) => break,
                                _ => {}
                            }
                        }
                        depth -= 1; // consumed the txbxContent end tag
                    }
                    // VML shape types
                    "shape" | "rect" | "oval" | "roundrect" | "line" => {
                        shape_type = Some(match local.as_str() {
                            "shape" => "rect", // generic shape defaults to rect
                            "roundrect" => "roundRect",
                            other => other,
                        }.to_string());
                        // Parse style attribute for width/height
                        for attr in e.attributes().flatten() {
                            let key = local_name(attr.key.as_ref());
                            let val = String::from_utf8_lossy(&attr.value).to_string();
                            match key.as_str() {
                                "style" => {
                                    // Parse CSS-like style: "width:200pt;height:100pt"
                                    for part in val.split(';') {
                                        let part = part.trim();
                                        if let Some(w) = part.strip_prefix("width:") {
                                            width = parse_css_length(w.trim());
                                        } else if let Some(h) = part.strip_prefix("height:") {
                                            height = parse_css_length(h.trim());
                                        } else if let Some(anchor) = part.strip_prefix("v-text-anchor:") {
                                            v_text_anchor = Some(anchor.trim().to_string());
                                        }
                                    }
                                }
                                "fillcolor" => fill_color = Some(val.trim_start_matches('#').to_string()),
                                "strokecolor" => stroke_color_val = Some(val.trim_start_matches('#').to_string()),
                                "strokeweight" => stroke_width_val = parse_css_length_opt(&val),
                                "stroked" => { if val == "f" || val == "false" { no_stroke = true; } }
                                _ => {}
                            }
                        }
                    }
                    _ => {}
                }
            }
            Event::Empty(e) => {
                let local = local_name(e.name().as_ref());
                match local.as_str() {
                    "imagedata" => {
                        for attr in e.attributes().flatten() {
                            let key = std::str::from_utf8(attr.key.as_ref()).unwrap_or("");
                            if key == "r:id" || key.ends_with(":id") {
                                rel_id = Some(String::from_utf8_lossy(&attr.value).to_string());
                            }
                        }
                    }
                    "fill" => {
                        for attr in e.attributes().flatten() {
                            if local_name(attr.key.as_ref()) == "color" {
                                fill_color = Some(String::from_utf8_lossy(&attr.value).trim_start_matches('#').to_string());
                            }
                        }
                    }
                    "stroke" => {
                        for attr in e.attributes().flatten() {
                            let key = local_name(attr.key.as_ref());
                            let val = String::from_utf8_lossy(&attr.value).to_string();
                            match key.as_str() {
                                "color" => stroke_color_val = Some(val.trim_start_matches('#').to_string()),
                                "weight" => stroke_width_val = parse_css_length_opt(&val),
                                "on" => { if val == "f" || val == "false" { no_stroke = true; } }
                                _ => {}
                            }
                        }
                    }
                    _ => {}
                }
            }
            Event::End(e) => {
                let local = local_name(e.name().as_ref());
                if local == "pict" && depth == 0 {
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

    // Build image if we have a blip reference
    let image = if let Some(rid) = rel_id {
        let data = ctx.media.get(&rid).cloned().unwrap_or_default();
        let content_type = ctx.media_types.get(&rid).cloned();
        Some(Image {
            data,
            width,
            height,
            alt_text: None,
            content_type,
            position: None,
            wrap_type: None,
            crop: None,
            anchor_block_index: 0,
        })
    } else {
        None
    };

    let shape = shape_type.as_ref().map(|st| Shape {
        shape_type: st.clone(),
        width,
        height,
        position: None,
        fill: fill_color.clone(),
        stroke_color: if no_stroke { None } else { stroke_color_val },
        stroke_width: if no_stroke { None } else { stroke_width_val },
        text_blocks,
        rotation: None,
        gradient_stops: Vec::new(),
        gradient_angle: None,
        anchor_block_index: 0,
        v_text_anchor,
    });

    Ok(DrawingResult { image, shape, text_box: None })
}

/// Parse CSS-like length value (e.g. "200pt", "2in", "100.5px")
fn parse_css_length(s: &str) -> f32 {
    let s = s.trim();
    if let Some(v) = s.strip_suffix("pt") {
        v.trim().parse().unwrap_or(0.0)
    } else if let Some(v) = s.strip_suffix("in") {
        v.trim().parse::<f32>().unwrap_or(0.0) * 72.0
    } else if let Some(v) = s.strip_suffix("cm") {
        v.trim().parse::<f32>().unwrap_or(0.0) * 28.3465
    } else if let Some(v) = s.strip_suffix("mm") {
        v.trim().parse::<f32>().unwrap_or(0.0) * 2.83465
    } else if let Some(v) = s.strip_suffix("px") {
        v.trim().parse::<f32>().unwrap_or(0.0) * 0.75 // 96dpi → 72pt
    } else {
        s.parse().unwrap_or(0.0)
    }
}

fn parse_css_length_opt(s: &str) -> Option<f32> {
    let v = parse_css_length(s);
    if v > 0.0 { Some(v) } else { None }
}

/// Parse w:object (OLE embedded object) — extract preview image from VML shape inside
fn parse_ole_object(reader: &mut Reader<&[u8]>, ctx: &ParseContext) -> Result<DrawingResult, ParseError> {
    let mut rel_id: Option<String> = None;
    let mut width: f32 = 0.0;
    let mut height: f32 = 0.0;
    let mut depth = 0;

    loop {
        match reader.read_event()? {
            Event::Start(e) => {
                let local = local_name(e.name().as_ref());
                depth += 1;
                match local.as_str() {
                    // VML shape inside OLE object — parse style for dimensions
                    "shape" | "rect" => {
                        for attr in e.attributes().flatten() {
                            let key = local_name(attr.key.as_ref());
                            if key == "style" {
                                let val = String::from_utf8_lossy(&attr.value);
                                for part in val.split(';') {
                                    let part = part.trim();
                                    if let Some(w) = part.strip_prefix("width:") {
                                        width = parse_css_length(w.trim());
                                    } else if let Some(h) = part.strip_prefix("height:") {
                                        height = parse_css_length(h.trim());
                                    }
                                }
                            }
                        }
                    }
                    _ => {}
                }
            }
            Event::Empty(e) => {
                let local = local_name(e.name().as_ref());
                match local.as_str() {
                    // v:imagedata — the preview image of the OLE object
                    "imagedata" => {
                        for attr in e.attributes().flatten() {
                            let key = std::str::from_utf8(attr.key.as_ref()).unwrap_or("");
                            if key == "r:id" || key.ends_with(":id") {
                                rel_id = Some(String::from_utf8_lossy(&attr.value).to_string());
                            }
                        }
                    }
                    // OLEObject element — skip gracefully
                    "OLEObject" => {}
                    _ => {}
                }
            }
            Event::End(e) => {
                let local = local_name(e.name().as_ref());
                if local == "object" && depth == 0 {
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

    // Build image from the preview if available
    let image = if let Some(rid) = rel_id {
        let data = ctx.media.get(&rid).cloned().unwrap_or_default();
        let content_type = ctx.media_types.get(&rid).cloned();
        Some(Image {
            data,
            width,
            height,
            alt_text: Some("OLE Object".to_string()),
            content_type,
            position: None,
            wrap_type: None,
            crop: None,
            anchor_block_index: 0,
        })
    } else {
        None
    };

    Ok(DrawingResult { image, shape: None, text_box: None })
}

/// Parse mc:AlternateContent — prefer mc:Choice (DrawingML), fall back to mc:Fallback (VML)
fn parse_alternate_content(reader: &mut Reader<&[u8]>, ctx: &ParseContext, styles: &StyleSheet) -> Result<Option<DrawingResult>, ParseError> {
    let mut result: Option<DrawingResult> = None;
    let mut depth = 0;
    let mut in_choice = false;
    let mut in_fallback = false;

    loop {
        match reader.read_event()? {
            Event::Start(e) => {
                let local = local_name(e.name().as_ref());
                match local.as_str() {
                    "Choice" if depth == 0 => {
                        in_choice = true;
                        depth += 1;
                    }
                    "Fallback" if depth == 0 => {
                        in_fallback = true;
                        depth += 1;
                    }
                    "drawing" if in_choice && depth == 1 => {
                        let dr = parse_drawing(reader, ctx, styles)?;
                        // Only keep if it produced something useful (image, shape, or text box)
                        if result.is_none() && dr.has_content() {
                            result = Some(dr);
                        }
                    }
                    "pict" if (in_choice || in_fallback) && depth == 1 && result.is_none() => {
                        let dr = parse_vml_pict(reader, ctx, styles)?;
                        if dr.has_content() {
                            result = Some(dr);
                        }
                    }
                    _ => {
                        depth += 1;
                    }
                }
            }
            Event::End(e) => {
                let local = local_name(e.name().as_ref());
                if local == "AlternateContent" && depth == 0 {
                    break;
                }
                if local == "Choice" && in_choice {
                    in_choice = false;
                }
                if local == "Fallback" && in_fallback {
                    in_fallback = false;
                }
                if depth > 0 {
                    depth -= 1;
                }
            }
            Event::Eof => break,
            _ => {}
        }
    }

    Ok(result)
}

/// Parse OMML math element (m:oMath or m:oMathPara) into a text representation
fn parse_omml(reader: &mut Reader<&[u8]>, end_tag: &str) -> Result<String, ParseError> {
    let mut result = String::new();
    let mut depth = 0;
    let mut in_text = false;
    // Track context for proper rendering
    let mut context_stack: Vec<String> = Vec::new();

    loop {
        match reader.read_event()? {
            Event::Start(e) => {
                let local = local_name(e.name().as_ref());
                depth += 1;
                match local.as_str() {
                    "t" => in_text = true,
                    "f" => context_stack.push("frac".to_string()),
                    "rad" => {
                        result.push('\u{221A}'); // √
                        context_stack.push("rad".to_string());
                    }
                    "sSup" => context_stack.push("sup".to_string()),
                    "sSub" => context_stack.push("sub".to_string()),
                    "d" => {
                        // Delimiter (parentheses)
                        let mut beg = '(';
                        for attr in e.attributes().flatten() {
                            if local_name(attr.key.as_ref()) == "begChr" {
                                let v = String::from_utf8_lossy(&attr.value);
                                beg = v.chars().next().unwrap_or('(');
                            }
                        }
                        result.push(beg);
                        context_stack.push("delim".to_string());
                    }
                    "nary" => {
                        // N-ary (summation, product, integral)
                        let mut chr = '\u{2211}'; // default: summation ∑
                        for attr in e.attributes().flatten() {
                            if local_name(attr.key.as_ref()) == "chr" {
                                let v = String::from_utf8_lossy(&attr.value);
                                chr = v.chars().next().unwrap_or('\u{2211}');
                            }
                        }
                        result.push(chr);
                        context_stack.push("nary".to_string());
                    }
                    "num" if context_stack.last().map_or(false, |c| c == "frac") => {
                        // Numerator of fraction — will add / separator after
                    }
                    "den" if context_stack.last().map_or(false, |c| c == "frac") => {
                        result.push('/');
                    }
                    "sup" if context_stack.last().map_or(false, |c| c == "sup") => {
                        result.push('^');
                    }
                    "sub" if context_stack.last().map_or(false, |c| c == "sub") => {
                        result.push('_');
                    }
                    _ => {}
                }
            }
            Event::Text(e) => {
                if in_text {
                    let content = e.unescape().unwrap_or_default();
                    result.push_str(&content);
                }
            }
            Event::End(e) => {
                let local = local_name(e.name().as_ref());
                if local == "t" {
                    in_text = false;
                }
                if local == end_tag && depth == 0 {
                    break;
                }
                match local.as_str() {
                    "f" | "rad" | "sSup" | "sSub" | "nary" => {
                        context_stack.pop();
                    }
                    "d" => {
                        context_stack.pop();
                        result.push(')');
                    }
                    _ => {}
                }
                if depth > 0 {
                    depth -= 1;
                }
            }
            Event::Empty(e) => {
                let local = local_name(e.name().as_ref());
                match local.as_str() {
                    // m:dPr begChr/endChr as empty element
                    "begChr" => {
                        for attr in e.attributes().flatten() {
                            if local_name(attr.key.as_ref()) == "val" {
                                // We already pushed '(' — if different, fix it
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

    Ok(result)
}

/// Parse w:rPr (run properties)
fn parse_run_properties(reader: &mut Reader<&[u8]>, ctx: &ParseContext, styles: &StyleSheet) -> Result<RunStyle, ParseError> {
    let mut style = RunStyle::default();
    let mut depth = 0;
    let mut rstyle_id: Option<String> = None;

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
                        } else if key == "eastAsia" {
                            style.font_family_east_asia =
                                Some(String::from_utf8_lossy(&attr.value).to_string());
                            style.has_explicit_east_asia = true;
                        } else if key == "asciiTheme" || key == "hAnsiTheme" {
                            if style.font_family.is_none() {
                                let val = String::from_utf8_lossy(&attr.value);
                                let font = super::styles::resolve_theme_font_pub(&val, &ctx.theme);
                                if let Some(f) = font {
                                    style.font_family = Some(f);
                                }
                            }
                        } else if key == "eastAsiaTheme" {
                            if style.font_family_east_asia.is_none() {
                                let val = String::from_utf8_lossy(&attr.value);
                                let font = super::styles::resolve_theme_font_pub(&val, &ctx.theme);
                                if let Some(f) = font {
                                    style.font_family_east_asia = Some(f);
                                }
                            }
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
                    "strike" => style.strikethrough = true,
                    "dstrike" => {
                        style.strikethrough = true;
                        style.double_strikethrough = true;
                    }
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
                            } else if key == "eastAsia" {
                                style.font_family_east_asia =
                                    Some(String::from_utf8_lossy(&attr.value).to_string());
                                style.has_explicit_east_asia = true;
                            } else if key == "asciiTheme" || key == "hAnsiTheme" {
                                if style.font_family.is_none() {
                                    let val = String::from_utf8_lossy(&attr.value);
                                    let font = super::styles::resolve_theme_font_pub(&val, &ctx.theme);
                                    if let Some(f) = font {
                                        style.font_family = Some(f);
                                    }
                                }
                            } else if key == "eastAsiaTheme" {
                                if style.font_family_east_asia.is_none() {
                                    let val = String::from_utf8_lossy(&attr.value);
                                    let font = super::styles::resolve_theme_font_pub(&val, &ctx.theme);
                                    if let Some(f) = font {
                                        style.font_family_east_asia = Some(f);
                                    }
                                }
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
                    "shd" => {
                        for attr in e.attributes().flatten() {
                            let key = local_name(attr.key.as_ref());
                            if key == "fill" {
                                let val = String::from_utf8_lossy(&attr.value).to_string();
                                if val != "auto" {
                                    style.shading = Some(val);
                                }
                            }
                        }
                    }
                    "rtl" => {
                        style.rtl = true;
                    }
                    "vanish" | "webHidden" => {
                        style.vanish = true;
                    }
                    "outline" => {
                        style.outline = true;
                    }
                    "shadow" => {
                        style.shadow = true;
                    }
                    "emboss" => {
                        style.emboss = true;
                    }
                    "imprint" => {
                        style.imprint = true;
                    }
                    "szCs" => {
                        for attr in e.attributes().flatten() {
                            if local_name(attr.key.as_ref()) == "val" {
                                let val = String::from_utf8_lossy(&attr.value);
                                style.font_size_cs = val.parse::<f32>().ok().map(|v| v / 2.0);
                            }
                        }
                    }
                    "bCs" => {
                        style.bold_cs = true;
                    }
                    "iCs" => {
                        style.italic_cs = true;
                    }
                    "kern" => {
                        for attr in e.attributes().flatten() {
                            if local_name(attr.key.as_ref()) == "val" {
                                let val = String::from_utf8_lossy(&attr.value);
                                style.kern = val.parse::<f32>().ok().map(|v| v / 2.0);
                            }
                        }
                    }
                    "fitText" => {
                        for attr in e.attributes().flatten() {
                            if local_name(attr.key.as_ref()) == "val" {
                                let val = String::from_utf8_lossy(&attr.value);
                                style.fit_text = val.parse::<f32>().ok().map(|v| v / 20.0);
                            }
                        }
                    }
                    "rStyle" => {
                        for attr in e.attributes().flatten() {
                            if local_name(attr.key.as_ref()) == "val" {
                                rstyle_id = Some(String::from_utf8_lossy(&attr.value).to_string());
                            }
                        }
                    }
                    "eastAsianLayout" => {
                        for attr in e.attributes().flatten() {
                            let key = local_name(attr.key.as_ref());
                            match key.as_str() {
                                "combine" => {
                                    let val = String::from_utf8_lossy(&attr.value);
                                    style.combine = val.as_ref() != "0" && val.as_ref() != "false";
                                }
                                "vert" => {
                                    let val = String::from_utf8_lossy(&attr.value);
                                    style.vert_in_horz = val.as_ref() != "0" && val.as_ref() != "false";
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

    // Apply character style (rStyle): matches Word output
    // char style properties are base, direct formatting overrides
    if let Some(ref id) = rstyle_id {
        if let Some(sdef) = styles.styles.get(id) {
            if let Some(ref rs) = sdef.paragraph.default_run_style {
                super::styles::merge_run_style(&mut style, rs);
            }
        }
    }

    Ok(style)
}

/// Parse a w:tbl element (table)
fn parse_table(reader: &mut Reader<&[u8]>, ctx: &ParseContext, styles: &StyleSheet) -> Result<Table, ParseError> {
    let mut rows = Vec::new();
    let mut style = TableStyle::default();
    let mut grid_columns = Vec::new();
    let mut depth = 0;

    loop {
        match reader.read_event()? {
            Event::Start(e) => {
                let local = local_name(e.name().as_ref());
                match local.as_str() {
                    "tblPr" if depth == 0 => {
                        style = parse_table_properties(reader)?;
                    }
                    "tblGrid" if depth == 0 => {
                        grid_columns = parse_table_grid(reader)?;
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

    // ECMA-376: Apply table style borders if table doesn't have explicit tblBorders.
    // Priority: tcBorders > tblStylePr > tblStyle tblBorders > direct tblBorders
    if !style.border {
        if let Some(ref style_id) = style.style_id {
            if let Some(tbl_style) = styles.table_styles.get(style_id) {
                if tbl_style.border {
                    style.border = true;
                    if tbl_style.has_inside_h {
                        style.has_inside_h = true;
                    }
                    if style.border_color.is_none() {
                        style.border_color = tbl_style.border_color.clone();
                    }
                    if style.border_width.is_none() {
                        style.border_width = tbl_style.border_width;
                    }
                    if style.border_style.is_none() {
                        style.border_style = tbl_style.border_style.clone();
                    }
                }
                // Inherit default cell margins from table style
                if style.default_cell_margins.is_none() {
                    style.default_cell_margins = tbl_style.default_cell_margins.clone();
                }
                // Inherit paragraph properties from table style
                if style.para_style.is_none() {
                    style.para_style = tbl_style.para_style.clone();
                }
            }
        }
    }

    Ok(Table { rows, style, grid_columns })
}

/// Parse w:tblGrid element — extract gridCol widths (twips → points)
fn parse_table_grid(reader: &mut Reader<&[u8]>) -> Result<Vec<f32>, ParseError> {
    let mut columns = Vec::new();
    loop {
        match reader.read_event()? {
            Event::Empty(e) | Event::Start(e) => {
                let local = local_name(e.name().as_ref());
                if local == "gridCol" {
                    for attr in e.attributes().flatten() {
                        let key = local_name(attr.key.as_ref());
                        if key == "w" {
                            if let Ok(val) = std::str::from_utf8(&attr.value) {
                                if let Ok(twips) = val.parse::<f32>() {
                                    columns.push(twips / 20.0); // twips to points
                                }
                            }
                        }
                    }
                }
            }
            Event::End(e) => {
                let local = local_name(e.name().as_ref());
                if local == "tblGrid" {
                    break;
                }
            }
            Event::Eof => break,
            _ => {}
        }
    }
    Ok(columns)
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
                    // Don't set border=true here; individual border elements check val!=none
                    in_borders = true;
                } else if local == "tblCellMar" {
                    // Parse default cell margins
                    let mut margins = CellMargins { top: None, bottom: None, left: None, right: None };
                    loop {
                        match reader.read_event() {
                            Ok(Event::Empty(me)) => {
                                let ml = local_name(me.name().as_ref());
                                let mut w_val: Option<f32> = None;
                                for attr in me.attributes().flatten() {
                                    if local_name(attr.key.as_ref()) == "w" {
                                        w_val = String::from_utf8_lossy(&attr.value).parse::<f32>().ok().map(|v| v / 20.0);
                                    }
                                }
                                match ml.as_str() {
                                    "top" => margins.top = w_val,
                                    "bottom" => margins.bottom = w_val,
                                    "left" | "start" => margins.left = w_val,
                                    "right" | "end" => margins.right = w_val,
                                    _ => {}
                                }
                            }
                            Ok(Event::End(ee)) => {
                                if local_name(ee.name().as_ref()) == "tblCellMar" { break; }
                            }
                            Ok(Event::Eof) => break,
                            _ => {}
                        }
                    }
                    style.default_cell_margins = Some(margins);
                    // Don't increment depth — we consumed the End event
                    continue;
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
                        // Check if border val is "none" or "nil" — if so, don't set border=true
                        let mut is_none = false;
                        let mut border_color_val = None;
                        let mut border_sz_val = None;
                        for attr in e.attributes().flatten() {
                            let key = local_name(attr.key.as_ref());
                            let val = String::from_utf8_lossy(&attr.value);
                            match key.as_str() {
                                "val" => {
                                    if val == "none" || val == "nil" {
                                        is_none = true;
                                    }
                                }
                                "color" => {
                                    if val != "auto" {
                                        border_color_val = Some(val.to_string());
                                    }
                                }
                                "sz" => {
                                    border_sz_val = val.parse::<f32>().ok().map(|v| v / 8.0);
                                }
                                _ => {}
                            }
                        }
                        if !is_none {
                            style.border = true;
                            if local == "insideH" {
                                style.has_inside_h = true;
                            }
                            if style.border_color.is_none() {
                                style.border_color = border_color_val;
                            }
                            if style.border_width.is_none() {
                                style.border_width = border_sz_val;
                            }
                            style.border_style = Some("single".to_string());
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
                    "tblStyleRowBandSize" => {
                        for attr in e.attributes().flatten() {
                            if local_name(attr.key.as_ref()) == "val" {
                                if let Some(ref mut look) = style.tbl_look {
                                    look.row_band_size = String::from_utf8_lossy(&attr.value).parse().unwrap_or(1);
                                }
                            }
                        }
                    }
                    "tblStyleColBandSize" => {
                        for attr in e.attributes().flatten() {
                            if local_name(attr.key.as_ref()) == "val" {
                                if let Some(ref mut look) = style.tbl_look {
                                    look.col_band_size = String::from_utf8_lossy(&attr.value).parse().unwrap_or(1);
                                }
                            }
                        }
                    }
                    "tblInd" => {
                        for attr in e.attributes().flatten() {
                            let key = local_name(attr.key.as_ref());
                            let val = String::from_utf8_lossy(&attr.value);
                            if key == "w" {
                                style.indent = val.parse::<f32>().ok().map(|v| v / 20.0);
                            }
                        }
                    }
                    "tblCellSpacing" => {
                        for attr in e.attributes().flatten() {
                            let key = local_name(attr.key.as_ref());
                            let val = String::from_utf8_lossy(&attr.value);
                            if key == "w" {
                                style.cell_spacing = val.parse::<f32>().ok().map(|v| v / 20.0);
                            }
                        }
                    }
                    "tblLook" => {
                        let mut look = TableLook::default();
                        for attr in e.attributes().flatten() {
                            let key = local_name(attr.key.as_ref());
                            let val = String::from_utf8_lossy(&attr.value);
                            match key.as_str() {
                                "firstRow" => look.first_row = val.as_ref() == "1",
                                "lastRow" => look.last_row = val.as_ref() == "1",
                                "firstColumn" => look.first_column = val.as_ref() == "1",
                                "lastColumn" => look.last_column = val.as_ref() == "1",
                                "noHBand" => look.banded_rows = val.as_ref() != "1",
                                "noVBand" => look.banded_columns = val.as_ref() != "1",
                                "val" => {
                                    // Hex bitmask fallback (e.g. "04A0")
                                    if let Ok(v) = u32::from_str_radix(&val, 16) {
                                        look.first_row = v & 0x0020 != 0;
                                        look.last_row = v & 0x0040 != 0;
                                        look.first_column = v & 0x0080 != 0;
                                        look.last_column = v & 0x0100 != 0;
                                        look.banded_rows = v & 0x0200 == 0; // noHBand inverted
                                        look.banded_columns = v & 0x0400 == 0; // noVBand inverted
                                    }
                                }
                                _ => {}
                            }
                        }
                        style.tbl_look = Some(look);
                    }
                    "tblLayout" => {
                        for attr in e.attributes().flatten() {
                            if local_name(attr.key.as_ref()) == "type" {
                                style.layout = Some(String::from_utf8_lossy(&attr.value).to_string());
                            }
                        }
                    }
                    "tblpPr" => {
                        let mut tp = TablePosition {
                            x: 0.0, y: 0.0,
                            h_anchor: None, v_anchor: None, h_align: None,
                            left_from_text: 0.0, right_from_text: 0.0,
                            top_from_text: 0.0, bottom_from_text: 0.0,
                        };
                        for attr in e.attributes().flatten() {
                            let key = local_name(attr.key.as_ref());
                            let val = String::from_utf8_lossy(&attr.value);
                            match key.as_str() {
                                "tblpX" => tp.x = val.parse::<f32>().unwrap_or(0.0) / 20.0,
                                "tblpY" => tp.y = val.parse::<f32>().unwrap_or(0.0) / 20.0,
                                "tblpXSpec" => tp.h_align = Some(val.to_string()),
                                "horzAnchor" => tp.h_anchor = Some(val.to_string()),
                                "vertAnchor" => tp.v_anchor = Some(val.to_string()),
                                "leftFromText" => tp.left_from_text = val.parse::<f32>().unwrap_or(0.0) / 20.0,
                                "rightFromText" => tp.right_from_text = val.parse::<f32>().unwrap_or(0.0) / 20.0,
                                "topFromText" => tp.top_from_text = val.parse::<f32>().unwrap_or(0.0) / 20.0,
                                "bottomFromText" => tp.bottom_from_text = val.parse::<f32>().unwrap_or(0.0) / 20.0,
                                _ => {}
                            }
                        }
                        style.position = Some(tp);
                    }
                    "top" | "bottom" | "left" | "right" | "start" | "end"
                        if !in_borders =>
                    {
                        // tblCellMar children (when not inside tblBorders)
                        // These appear inside <w:tblCellMar> which is a Start element
                        // but we handle them as Empty within depth tracking
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
    let mut height_rule: Option<String> = None;
    let mut header = false;
    let mut cant_split = false;
    let mut grid_before: u32 = 0;
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
                match local.as_str() {
                    "trHeight" => {
                        for attr in e.attributes().flatten() {
                            let key = local_name(attr.key.as_ref());
                            if key == "val" {
                                let val = String::from_utf8_lossy(&attr.value);
                                height = val.parse::<f32>().ok().map(|v| v / 20.0);
                            }
                            if key == "hRule" {
                                height_rule = Some(String::from_utf8_lossy(&attr.value).to_string());
                            }
                        }
                    }
                    "tblHeader" => { header = true; }
                    "cantSplit" => { cant_split = true; }
                    "gridBefore" => {
                        for attr in e.attributes().flatten() {
                            if local_name(attr.key.as_ref()) == "val" {
                                grid_before = String::from_utf8_lossy(&attr.value).parse().unwrap_or(0);
                            }
                        }
                    }
                    _ => {}
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

    Ok(TableRow { cells, height, height_rule, header, cant_split, grid_before })
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
                        let pr = parse_paragraph(reader, ctx, styles)?;
                        blocks.push(Block::Paragraph(pr.paragraph));
                    }
                    "tbl" if depth == 0 => {
                        let table = parse_table(reader, ctx, styles)?;
                        blocks.push(Block::Table(table));
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

/// A header/footer reference with its type
#[derive(Debug, Clone)]
struct HdrFtrRef {
    rel_id: String,
    ref_type: String, // "default", "first", "even"
}

/// Parsed section properties
struct SectionProperties {
    page_size: PageSize,
    margin: Margin,
    /// Document grid line pitch in points (from w:docGrid w:linePitch, twips/20)
    grid_line_pitch: Option<f32>,
    /// Character grid pitch in points (for linesAndChars mode)
    grid_char_pitch: Option<f32>,
    /// docGrid exists but has no type attribute
    doc_grid_no_type: bool,
    /// Reference IDs for header parts (with type)
    header_refs: Vec<HdrFtrRef>,
    /// Reference IDs for footer parts (with type)
    footer_refs: Vec<HdrFtrRef>,
    /// Column layout
    columns: Option<ColumnLayout>,
    /// Whether this section has a different first page header/footer
    title_pg: bool,
    /// Section break type: "nextPage" (default), "continuous", "evenPage", "oddPage"
    section_type: Option<String>,
    /// Page number format (e.g. "decimal", "lowerRoman", "upperRoman", "lowerLetter", "upperLetter")
    page_number_format: Option<String>,
    /// Starting page number for this section
    page_number_start: Option<u32>,
    /// Page borders
    page_borders: Option<PageBorders>,
    /// Header distance from page top edge (w:pgMar header attr, in points)
    header_distance: Option<f32>,
    /// Footer distance from page bottom edge (w:pgMar footer attr, in points)
    footer_distance: Option<f32>,
}

/// Parse w:sectPr (section properties - page size, margins, document grid)
fn parse_section_properties(
    reader: &mut Reader<&[u8]>,
) -> Result<SectionProperties, ParseError> {
    let mut page_size = PageSize::default();
    let mut margin = Margin::default();
    let mut grid_line_pitch: Option<f32> = None;
    let mut grid_char_pitch: Option<f32> = None;
    let mut doc_grid_no_type = false;
    let mut header_refs: Vec<HdrFtrRef> = Vec::new();
    let mut footer_refs: Vec<HdrFtrRef> = Vec::new();
    let mut columns: Option<ColumnLayout> = None;
    let mut title_pg = false;
    let mut section_type: Option<String> = None;
    let mut page_number_format: Option<String> = None;
    let mut page_number_start: Option<u32> = None;
    let mut page_borders: Option<PageBorders> = None;
    let mut header_distance: Option<f32> = None;
    let mut footer_distance: Option<f32> = None;
    let mut depth = 0;

    loop {
        match reader.read_event()? {
            Event::Start(e) => {
                let local = local_name(e.name().as_ref());
                if local == "pgBorders" && depth == 0 {
                    // Parse page borders - child elements: top, bottom, left, right
                    let mut pb = PageBorders { top: None, bottom: None, left: None, right: None };
                    loop {
                        match reader.read_event() {
                            Ok(Event::Empty(be)) => {
                                let bl = local_name(be.name().as_ref());
                                let mut bdr_style = String::new();
                                let mut bdr_width = 0.0f32;
                                let mut bdr_color: Option<String> = None;
                                let mut bdr_space = 0.0f32;
                                for attr in be.attributes().flatten() {
                                    let key = local_name(attr.key.as_ref());
                                    let val = String::from_utf8_lossy(&attr.value);
                                    match key.as_str() {
                                        "val" => bdr_style = val.to_string(),
                                        "sz" => { bdr_width = val.parse::<f32>().unwrap_or(0.0) / 8.0; }
                                        "color" => {
                                            let c = val.to_string();
                                            if c != "auto" { bdr_color = Some(c); }
                                        }
                                        "space" => { bdr_space = val.parse::<f32>().unwrap_or(0.0); }
                                        _ => {}
                                    }
                                }
                                if bdr_style != "none" && bdr_style != "nil" && bdr_width > 0.0 {
                                    let def = BorderDef { style: bdr_style, width: bdr_width, color: bdr_color, space: bdr_space };
                                    match bl.as_str() {
                                        "top" => pb.top = Some(def),
                                        "bottom" => pb.bottom = Some(def),
                                        "left" => pb.left = Some(def),
                                        "right" => pb.right = Some(def),
                                        _ => {}
                                    }
                                }
                            }
                            Ok(Event::End(ee)) => {
                                if local_name(ee.name().as_ref()) == "pgBorders" { break; }
                            }
                            Ok(Event::Eof) => break,
                            _ => {}
                        }
                    }
                    if pb.top.is_some() || pb.bottom.is_some() || pb.left.is_some() || pb.right.is_some() {
                        page_borders = Some(pb);
                    }
                } else if local == "cols" && depth == 0 {
                    // w:cols as Start element — has child w:col elements
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
                    // Parse child w:col elements
                    let mut col_defs = Vec::new();
                    loop {
                        match reader.read_event() {
                            Ok(Event::Empty(ce)) => {
                                let cl = local_name(ce.name().as_ref());
                                if cl == "col" {
                                    let mut col_w = 0.0f32;
                                    let mut col_space: Option<f32> = None;
                                    for attr in ce.attributes().flatten() {
                                        let key = local_name(attr.key.as_ref());
                                        let val = String::from_utf8_lossy(&attr.value);
                                        match key.as_str() {
                                            "w" => { col_w = val.parse::<f32>().unwrap_or(0.0) / 20.0; }
                                            "space" => { col_space = val.parse::<f32>().ok().map(|v| v / 20.0); }
                                            _ => {}
                                        }
                                    }
                                    col_defs.push(ColumnDef { width: col_w, space: col_space });
                                }
                            }
                            Ok(Event::End(ee)) => {
                                if local_name(ee.name().as_ref()) == "cols" {
                                    break;
                                }
                            }
                            Ok(Event::Eof) => break,
                            _ => {}
                        }
                    }
                    if num > 1 {
                        columns = Some(ColumnLayout { num, space, equal_width, columns: col_defs });
                    }
                } else {
                    depth += 1;
                }
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
                        // COM-confirmed (2026-04-03): Word rounds page margins
                        // to nearest 0.5pt (10 twips). round(twips/10)*0.5
                        let round_10tw = |tw: f32| -> f32 {
                            (tw / 10.0).round() * 0.5
                        };
                        let mut gutter = 0.0f32;
                        for attr in e.attributes().flatten() {
                            let key = local_name(attr.key.as_ref());
                            let val = String::from_utf8_lossy(&attr.value);
                            match key.as_str() {
                                "top" => {
                                    if let Ok(v) = val.parse::<f32>() {
                                        margin.top = round_10tw(v);
                                    }
                                }
                                "bottom" => {
                                    if let Ok(v) = val.parse::<f32>() {
                                        margin.bottom = round_10tw(v);
                                    }
                                }
                                "left" => {
                                    if let Ok(v) = val.parse::<f32>() {
                                        margin.left = round_10tw(v);
                                    }
                                }
                                "right" => {
                                    if let Ok(v) = val.parse::<f32>() {
                                        margin.right = round_10tw(v);
                                    }
                                }
                                "gutter" => {
                                    if let Ok(v) = val.parse::<f32>() {
                                        gutter = round_10tw(v);
                                    }
                                }
                                "header" => {
                                    if let Ok(v) = val.parse::<f32>() {
                                        header_distance = Some(round_10tw(v));
                                    }
                                }
                                "footer" => {
                                    if let Ok(v) = val.parse::<f32>() {
                                        footer_distance = Some(round_10tw(v));
                                    }
                                }
                                _ => {}
                            }
                        }
                        // Gutter adds to left margin (default) or top margin (gutterTop not implemented yet)
                        if gutter > 0.0 {
                            margin.left += gutter;
                        }
                    }
                    "docGrid" => {
                        let mut grid_type = String::new();
                        let mut line_pitch = 0u32;
                        // §17.6.5 (Round 15, COM-confirmed): w:charSpace is a SIGNED
                        // integer (can be negative for compressed grid). u32 silently
                        // dropped negative values to 0, causing wrong pitch on docs
                        // with charSpace<0.
                        let mut char_space: Option<i32> = None;
                        for attr in e.attributes().flatten() {
                            let key = local_name(attr.key.as_ref());
                            let val = String::from_utf8_lossy(&attr.value);
                            match key.as_str() {
                                "type" => grid_type = val.to_string(),
                                "linePitch" => {
                                    line_pitch = val.parse().unwrap_or(0);
                                }
                                "charSpace" => {
                                    char_space = Some(val.parse().unwrap_or(0));
                                }
                                _ => {}
                            }
                        }
                        // Only apply grid for "lines" or "linesAndChars" types
                        if (grid_type == "lines" || grid_type == "linesAndChars")
                            && line_pitch > 0
                        {
                            grid_line_pitch = Some(line_pitch as f32 / 20.0);
                        } else if grid_type.is_empty() && line_pitch > 0 {
                            // docGrid exists with linePitch but no type attribute
                            doc_grid_no_type = true;
                        }
                        // linesAndChars: compute character grid pitch
                        // COM-confirmed (2026-04-03): charGrid is active even without charSpace.
                        // Formula: raw_pitch = default_font_size + charSpace/4096
                        //          charsLine = floor(contentWidth / raw_pitch)
                        //          actual_pitch = contentWidth / charsLine
                        // charSpace unit: 1/4096 of a point (ECMA-376 §17.6.5)
                        if grid_type == "linesAndChars" {
                            // charGrid raw_pitch uses the document's default font size.
                            // This comes from Normal style's sz, or rPrDefault sz, or 10.5pt fallback.
                            // Stored in SectionProperties and resolved by the caller post-parse.
                            let default_font_size = 10.5_f32; // placeholder; overridden by caller
                            let char_space_pt = char_space.map(|cs| cs as f32 / 4096.0).unwrap_or(0.0);
                            let raw_pitch = default_font_size + char_space_pt;
                            let content_w = page_size.width - margin.left - margin.right;
                            if raw_pitch > 0.0 && content_w > 0.0 {
                                let chars_line = (content_w / raw_pitch).floor().max(1.0);
                                grid_char_pitch = Some(content_w / chars_line);
                            }
                        }
                    }
                    "headerReference" => {
                        let mut rel_id = String::new();
                        let mut ref_type = "default".to_string();
                        for attr in e.attributes().flatten() {
                            let key = local_name(attr.key.as_ref());
                            let val = String::from_utf8_lossy(&attr.value).to_string();
                            match key.as_str() {
                                "id" => rel_id = val,
                                "type" => ref_type = val,
                                _ => {}
                            }
                        }
                        if !rel_id.is_empty() {
                            header_refs.push(HdrFtrRef { rel_id, ref_type });
                        }
                    }
                    "footerReference" => {
                        let mut rel_id = String::new();
                        let mut ref_type = "default".to_string();
                        for attr in e.attributes().flatten() {
                            let key = local_name(attr.key.as_ref());
                            let val = String::from_utf8_lossy(&attr.value).to_string();
                            match key.as_str() {
                                "id" => rel_id = val,
                                "type" => ref_type = val,
                                _ => {}
                            }
                        }
                        if !rel_id.is_empty() {
                            footer_refs.push(HdrFtrRef { rel_id, ref_type });
                        }
                    }
                    "titlePg" => {
                        title_pg = true;
                    }
                    "type" => {
                        for attr in e.attributes().flatten() {
                            if local_name(attr.key.as_ref()) == "val" {
                                section_type = Some(String::from_utf8_lossy(&attr.value).to_string());
                            }
                        }
                    }
                    "pgNumType" => {
                        for attr in e.attributes().flatten() {
                            let key = local_name(attr.key.as_ref());
                            let val = String::from_utf8_lossy(&attr.value);
                            match key.as_str() {
                                "fmt" => {
                                    page_number_format = Some(val.to_string());
                                }
                                "start" => {
                                    page_number_start = val.parse::<u32>().ok();
                                }
                                _ => {}
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
                            columns = Some(ColumnLayout { num, space, equal_width, columns: Vec::new() });
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
        grid_char_pitch,
        doc_grid_no_type,
        header_refs,
        footer_refs,
        columns,
        title_pg,
        section_type,
        page_number_format,
        page_number_start,
        page_borders,
        header_distance,
        footer_distance,
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
                        let pr = parse_paragraph(&mut reader, ctx, styles)?;
                        blocks.push(Block::Paragraph(pr.paragraph));
                    }
                    "tbl" if in_root && depth == 0 => {
                        let table = parse_table(&mut reader, ctx, styles)?;
                        blocks.push(Block::Table(table));
                    }
                    "sdt" if in_root && depth == 0 => {
                        // Structured Document Tag: skip sdtPr, process sdtContent children
                        let mut sdt_depth = 1u32;
                        let mut in_sdt_content = false;
                        loop {
                            match reader.read_event()? {
                                Event::Start(se) => {
                                    let sl = local_name(se.name().as_ref());
                                    if sl == "sdtContent" && sdt_depth == 1 {
                                        in_sdt_content = true;
                                    } else if in_sdt_content && sl == "p" {
                                        let pr = parse_paragraph(&mut reader, ctx, styles)?;
                                        blocks.push(Block::Paragraph(pr.paragraph));
                                    } else if in_sdt_content && sl == "tbl" {
                                        let table = parse_table(&mut reader, ctx, styles)?;
                                        blocks.push(Block::Table(table));
                                    } else {
                                        sdt_depth += 1;
                                    }
                                }
                                Event::End(ee) => {
                                    let sl = local_name(ee.name().as_ref());
                                    if sl == "sdtContent" {
                                        in_sdt_content = false;
                                    } else if sl == "sdt" && sdt_depth == 1 {
                                        break;
                                    } else if sdt_depth > 1 {
                                        sdt_depth -= 1;
                                    }
                                }
                                Event::Eof => break,
                                _ => {}
                            }
                        }
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

/// Round 29: Rewrite footnoteReference / endnoteReference run text from
/// the OOXML id ("[2]") to the section-local sequential number ("[1]"),
/// using the supplied id→seq maps. Walks paragraphs and recursively into
/// table cells. Run.footnote_ref / Run.endnote_ref still hold the original
/// OOXML id so the rendering loop can look up the body by id.
fn renumber_note_refs(
    blocks: &mut Vec<Block>,
    fn_map: &std::collections::HashMap<u32, u32>,
    en_map: &std::collections::HashMap<u32, u32>,
) {
    for block in blocks.iter_mut() {
        match block {
            Block::Paragraph(para) => {
                for run in para.runs.iter_mut() {
                    if let Some(id) = run.footnote_ref {
                        if let Some(seq) = fn_map.get(&id) {
                            run.text = format!("{}", seq);
                        }
                    }
                    if let Some(id) = run.endnote_ref {
                        if let Some(seq) = en_map.get(&id) {
                            run.text = format!("{}", seq);
                        }
                    }
                }
            }
            Block::Table(table) => {
                for row in table.rows.iter_mut() {
                    for cell in row.cells.iter_mut() {
                        renumber_note_refs(&mut cell.blocks, fn_map, en_map);
                    }
                }
            }
            Block::Image(_) | Block::UnsupportedElement(_) => {}
        }
    }
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
            Block::Image(_) | Block::UnsupportedElement(_) => {}
        }
    }
}

/// Parse runs inside w:ins or w:del (tracked changes)
fn parse_tracked_change_runs(
    reader: &mut Reader<&[u8]>,
    ctx: &ParseContext,
    styles: &StyleSheet,
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
                    let (mut run, _dr) = parse_run(reader, ctx, styles, None)?;
                    // Matches Word output:
                    // - "delete" → strikethrough + red color
                    // - "insert" → underline + red color
                    if tc.change_type == "delete" {
                        run.style.strikethrough = true;
                        if run.style.color.is_none() {
                            run.style.color = Some("FF0000".to_string());
                        }
                    } else if tc.change_type == "insert" {
                        run.style.underline = true;
                        if run.style.color.is_none() {
                            run.style.color = Some("FF0000".to_string());
                        }
                    }
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
                        let pr = parse_paragraph(&mut reader, &note_ctx, &empty_styles)?;
                        let para = pr.paragraph;
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
