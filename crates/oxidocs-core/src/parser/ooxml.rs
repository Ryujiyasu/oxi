use std::io::{Cursor, Read};

use quick_xml::events::Event;
use quick_xml::reader::Reader;
use zip::ZipArchive;

use super::styles::parse_styles;
use super::ParseError;
use crate::ir::*;

pub struct OoxmlParser {
    archive: ZipArchive<Cursor<Vec<u8>>>,
}

impl OoxmlParser {
    pub fn new(data: &[u8]) -> Result<Self, ParseError> {
        let cursor = Cursor::new(data.to_vec());
        let archive = ZipArchive::new(cursor)?;
        Ok(Self { archive })
    }

    pub fn parse(mut self) -> Result<Document, ParseError> {
        let styles = self.parse_styles()?;
        let (blocks, sect_pr) = self.parse_document_xml()?;
        let metadata = self.parse_metadata();

        let (page_size, margin) = sect_pr.unwrap_or_else(|| {
            (PageSize::default(), Margin::default())
        });

        Ok(Document {
            pages: vec![Page {
                blocks,
                size: page_size,
                margin,
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

    fn parse_styles(&mut self) -> Result<StyleSheet, ParseError> {
        match self.read_part("word/styles.xml") {
            Ok(xml) => parse_styles(&xml),
            Err(ParseError::MissingPart(_)) => Ok(StyleSheet::default()),
            Err(e) => Err(e),
        }
    }

    fn parse_metadata(&self) -> DocumentMetadata {
        // TODO: Parse docProps/core.xml for title, author, etc.
        DocumentMetadata::default()
    }

    fn parse_document_xml(
        &mut self,
    ) -> Result<(Vec<Block>, Option<(PageSize, Margin)>), ParseError> {
        let xml = self.read_part("word/document.xml")?;
        parse_body(&xml)
    }
}

/// Parse the w:body content of document.xml
fn parse_body(xml: &str) -> Result<(Vec<Block>, Option<(PageSize, Margin)>), ParseError> {
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
                        let para = parse_paragraph(&mut reader)?;
                        blocks.push(Block::Paragraph(para));
                    }
                    "tbl" if in_body && depth == 0 => {
                        let table = parse_table(&mut reader)?;
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
fn parse_paragraph(reader: &mut Reader<&[u8]>) -> Result<Paragraph, ParseError> {
    let mut runs = Vec::new();
    let mut style = ParagraphStyle::default();
    let mut alignment = Alignment::default();
    let mut depth = 0;

    loop {
        match reader.read_event()? {
            Event::Start(e) => {
                let local = local_name(e.name().as_ref());
                match local.as_str() {
                    "pPr" if depth == 0 => {
                        let (s, a) = parse_paragraph_properties(reader)?;
                        style = s;
                        alignment = a;
                    }
                    "r" if depth == 0 => {
                        let run = parse_run(reader)?;
                        runs.push(run);
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

    Ok(Paragraph {
        runs,
        style,
        alignment,
    })
}

/// Parse w:pPr (paragraph properties)
fn parse_paragraph_properties(
    reader: &mut Reader<&[u8]>,
) -> Result<(ParagraphStyle, Alignment), ParseError> {
    let mut style = ParagraphStyle::default();
    let mut alignment = Alignment::default();
    let mut depth = 0;

    loop {
        match reader.read_event()? {
            Event::Start(e) => {
                let local = local_name(e.name().as_ref());
                match local.as_str() {
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
                                let val = String::from_utf8_lossy(&attr.value);
                                if val.starts_with("Heading") {
                                    style.heading_level =
                                        val.trim_start_matches("Heading").parse().ok();
                                }
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

    Ok((style, alignment))
}

/// Parse a w:r element (run)
fn parse_run(reader: &mut Reader<&[u8]>) -> Result<Run, ParseError> {
    let mut text = String::new();
    let mut style = RunStyle::default();
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

    Ok(Run { text, style })
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
fn parse_table(reader: &mut Reader<&[u8]>) -> Result<Table, ParseError> {
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
                        let row = parse_table_row(reader)?;
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
fn parse_table_row(reader: &mut Reader<&[u8]>) -> Result<TableRow, ParseError> {
    let mut cells = Vec::new();
    let mut depth = 0;

    loop {
        match reader.read_event()? {
            Event::Start(e) => {
                let local = local_name(e.name().as_ref());
                match local.as_str() {
                    "tc" if depth == 0 => {
                        let cell = parse_table_cell(reader)?;
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
fn parse_table_cell(reader: &mut Reader<&[u8]>) -> Result<TableCell, ParseError> {
    let mut blocks = Vec::new();
    let mut width = None;
    let mut depth = 0;

    loop {
        match reader.read_event()? {
            Event::Start(e) => {
                let local = local_name(e.name().as_ref());
                match local.as_str() {
                    "p" if depth == 0 => {
                        let para = parse_paragraph(reader)?;
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

/// Parse w:sectPr (section properties - page size and margins)
fn parse_section_properties(
    reader: &mut Reader<&[u8]>,
) -> Result<(PageSize, Margin), ParseError> {
    let mut page_size = PageSize::default();
    let mut margin = Margin::default();
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
                                // OOXML uses twips (1/20 of a point)
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

    Ok((page_size, margin))
}

/// Extract local name from a potentially namespaced XML tag
/// e.g., "w:body" -> "body", "body" -> "body"
fn local_name(name: &[u8]) -> String {
    let s = std::str::from_utf8(name).unwrap_or("");
    match s.rfind(':') {
        Some(pos) => s[pos + 1..].to_string(),
        None => s.to_string(),
    }
}
