pub mod editor;
pub mod font;
pub mod ir;
pub mod layout;
pub mod parser;

pub use editor::DocxEditor;
pub use ir::Document;
pub use parser::parse_docx;

/// Create a minimal blank .docx file as bytes.
/// The generated file contains a single empty paragraph and can be
/// parsed by `parse_docx` and edited by `DocxEditor`.
pub fn create_blank_docx() -> Vec<u8> {
    use std::io::{Cursor, Write};
    use zip::write::{SimpleFileOptions, ZipWriter};

    let buf = Cursor::new(Vec::new());
    let mut zip = ZipWriter::new(buf);
    let opts = SimpleFileOptions::default()
        .compression_method(zip::CompressionMethod::Deflated);

    // [Content_Types].xml
    zip.start_file("[Content_Types].xml", opts).unwrap();
    zip.write_all(br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
</Types>"#).unwrap();

    // _rels/.rels
    zip.start_file("_rels/.rels", opts).unwrap();
    zip.write_all(br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"#).unwrap();

    // word/_rels/document.xml.rels
    zip.start_file("word/_rels/document.xml.rels", opts).unwrap();
    zip.write_all(br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>"#).unwrap();

    // word/document.xml - single empty paragraph
    zip.start_file("word/document.xml", opts).unwrap();
    zip.write_all(br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    <w:p>
      <w:r><w:t></w:t></w:r>
    </w:p>
    <w:sectPr>
      <w:pgSz w:w="11906" w:h="16838"/>
      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/>
    </w:sectPr>
  </w:body>
</w:document>"#).unwrap();

    // word/styles.xml - Normal style with Calibri 11pt
    zip.start_file("word/styles.xml", opts).unwrap();
    zip.write_all(br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:docDefaults>
    <w:rPrDefault>
      <w:rPr>
        <w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/>
        <w:sz w:val="22"/>
      </w:rPr>
    </w:rPrDefault>
    <w:pPrDefault>
      <w:pPr>
        <w:spacing w:after="160" w:line="259" w:lineRule="auto"/>
      </w:pPr>
    </w:pPrDefault>
  </w:docDefaults>
  <w:style w:type="paragraph" w:styleId="Normal">
    <w:name w:val="Normal"/>
  </w:style>
</w:styles>"#).unwrap();

    zip.finish().unwrap().into_inner()
}

/// Content block for build_docx.
#[derive(Debug, Clone, serde::Deserialize)]
pub struct ContentBlock {
    #[serde(rename = "type")]
    pub block_type: String, // "paragraph" or "table"
    #[serde(default)]
    pub runs: Vec<ContentRun>,
    #[serde(default)]
    pub alignment: Option<String>,
    #[serde(default)]
    pub heading_level: Option<u8>,
    #[serde(default)]
    pub line_height: Option<f32>,
    // Table fields
    #[serde(default)]
    pub rows: Option<Vec<Vec<ContentCell>>>,
}

#[derive(Debug, Clone, serde::Deserialize)]
pub struct ContentRun {
    #[serde(default)]
    pub text: String,
    #[serde(default)]
    pub bold: Option<bool>,
    #[serde(default)]
    pub italic: Option<bool>,
    #[serde(default)]
    pub underline: Option<bool>,
    #[serde(default)]
    pub strikethrough: Option<bool>,
    #[serde(default)]
    pub font_family: Option<String>,
    #[serde(default)]
    pub font_size: Option<f32>,
    #[serde(default)]
    pub color: Option<String>,
}

#[derive(Debug, Clone, serde::Deserialize)]
pub struct ContentCell {
    #[serde(default)]
    pub text: String,
    #[serde(default)]
    pub bold: Option<bool>,
}

/// Build a .docx file from a list of content blocks.
/// This generates a complete docx with proper OOXML structure.
pub fn build_docx(blocks: &[ContentBlock]) -> Vec<u8> {
    use std::io::{Cursor, Write};
    use zip::write::{SimpleFileOptions, ZipWriter};

    let body_xml = generate_body_xml(blocks);

    // Build the ZIP
    let buf = Cursor::new(Vec::new());
    let mut zip = ZipWriter::new(buf);
    let opts = SimpleFileOptions::default()
        .compression_method(zip::CompressionMethod::Deflated);

    // [Content_Types].xml
    zip.start_file("[Content_Types].xml", opts).unwrap();
    zip.write_all(br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
</Types>"#).unwrap();

    zip.start_file("_rels/.rels", opts).unwrap();
    zip.write_all(br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"#).unwrap();

    zip.start_file("word/_rels/document.xml.rels", opts).unwrap();
    zip.write_all(br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>"#).unwrap();

    // document.xml with generated body
    zip.start_file("word/document.xml", opts).unwrap();
    let doc_xml = format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    {}
    <w:sectPr>
      <w:pgSz w:w="11906" w:h="16838"/>
      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/>
    </w:sectPr>
  </w:body>
</w:document>"#,
        body_xml
    );
    zip.write_all(doc_xml.as_bytes()).unwrap();

    // styles.xml
    zip.start_file("word/styles.xml", opts).unwrap();
    zip.write_all(br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:docDefaults>
    <w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="22"/></w:rPr></w:rPrDefault>
    <w:pPrDefault><w:pPr><w:spacing w:after="160" w:line="259" w:lineRule="auto"/></w:pPr></w:pPrDefault>
  </w:docDefaults>
  <w:style w:type="paragraph" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
  <w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:pPr><w:keepNext/><w:spacing w:before="240" w:after="0"/><w:outlineLvl w:val="0"/></w:pPr><w:rPr><w:b/><w:bCs/><w:sz w:val="32"/><w:szCs w:val="32"/></w:rPr></w:style>
  <w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/><w:basedOn w:val="Normal"/><w:pPr><w:keepNext/><w:spacing w:before="200" w:after="0"/><w:outlineLvl w:val="1"/></w:pPr><w:rPr><w:b/><w:bCs/><w:sz w:val="26"/><w:szCs w:val="26"/></w:rPr></w:style>
  <w:style w:type="paragraph" w:styleId="Heading3"><w:name w:val="heading 3"/><w:basedOn w:val="Normal"/><w:pPr><w:keepNext/><w:spacing w:before="200" w:after="0"/><w:outlineLvl w:val="2"/></w:pPr><w:rPr><w:b/><w:bCs/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr></w:style>
  <w:style w:type="paragraph" w:styleId="Heading4"><w:name w:val="heading 4"/><w:basedOn w:val="Normal"/><w:pPr><w:keepNext/><w:spacing w:before="200" w:after="0"/><w:outlineLvl w:val="3"/></w:pPr><w:rPr><w:b/><w:bCs/><w:i/><w:iCs/><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr></w:style>
  <w:style w:type="paragraph" w:styleId="Heading5"><w:name w:val="heading 5"/><w:basedOn w:val="Normal"/><w:pPr><w:keepNext/><w:spacing w:before="200" w:after="0"/><w:outlineLvl w:val="4"/></w:pPr><w:rPr><w:sz w:val="22"/><w:szCs w:val="22"/><w:color w:val="1F4D78"/></w:rPr></w:style>
  <w:style w:type="paragraph" w:styleId="Heading6"><w:name w:val="heading 6"/><w:basedOn w:val="Normal"/><w:pPr><w:keepNext/><w:spacing w:before="200" w:after="0"/><w:outlineLvl w:val="5"/></w:pPr><w:rPr><w:i/><w:iCs/><w:sz w:val="22"/><w:szCs w:val="22"/><w:color w:val="1F4D78"/></w:rPr></w:style>
</w:styles>"#).unwrap();

    zip.finish().unwrap().into_inner()
}

/// Build a .docx from content blocks, using a template docx for styles/theme/numbering.
/// Preserves the template's styles.xml, theme, numbering, fontTable, settings, etc.
/// Only replaces word/document.xml with generated content.
pub fn build_docx_with_template(blocks: &[ContentBlock], template: &[u8]) -> Vec<u8> {
    use std::io::{Cursor, Read, Write};
    use zip::write::{SimpleFileOptions, ZipWriter};
    use zip::ZipArchive;

    // Generate body XML (same logic as build_docx)
    let body_xml = generate_body_xml(blocks);

    let doc_xml = format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
            xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
            xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
            xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"
            xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">
  <w:body>
    {}
    <w:sectPr>
      <w:pgSz w:w="11906" w:h="16838"/>
      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/>
    </w:sectPr>
  </w:body>
</w:document>"#,
        body_xml
    );

    // Copy everything from template except document.xml
    let cursor = Cursor::new(template);
    let mut archive = match ZipArchive::new(cursor) {
        Ok(a) => a,
        Err(_) => return build_docx(blocks), // fallback
    };

    // Extract sectPr from original document.xml if present
    let original_sect_pr = if let Ok(mut entry) = archive.by_name("word/document.xml") {
        let mut xml = String::new();
        entry.read_to_string(&mut xml).ok();
        // Extract sectPr from original
        extract_sect_pr(&xml)
    } else {
        None
    };

    // If we got original sectPr, use it instead of the default
    let final_doc_xml = if let Some(ref sect) = original_sect_pr {
        let default_sect = r#"<w:sectPr>
      <w:pgSz w:w="11906" w:h="16838"/>
      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/>
    </w:sectPr>"#;
        doc_xml.replace(default_sect, sect)
    } else {
        doc_xml
    };

    let mut output = Vec::new();
    {
        let mut writer = ZipWriter::new(Cursor::new(&mut output));
        let opts = SimpleFileOptions::default()
            .compression_method(zip::CompressionMethod::Deflated);

        for i in 0..archive.len() {
            let mut entry = archive.by_index(i).unwrap();
            let name = entry.name().to_string();

            if name == "word/document.xml" {
                writer.start_file(&name, opts).unwrap();
                writer.write_all(final_doc_xml.as_bytes()).unwrap();
            } else {
                let entry_opts = SimpleFileOptions::default()
                    .compression_method(entry.compression());
                writer.start_file(&name, entry_opts).unwrap();
                let mut buf = Vec::new();
                entry.read_to_end(&mut buf).unwrap();
                writer.write_all(&buf).unwrap();
            }
        }
        writer.finish().unwrap();
    }
    output
}

/// Extract the <w:sectPr>...</w:sectPr> block from document.xml.
fn extract_sect_pr(xml: &str) -> Option<String> {
    // Simple extraction: find last <w:sectPr and corresponding </w:sectPr>
    let start = xml.rfind("<w:sectPr")?;
    let end_tag = "</w:sectPr>";
    let end = xml[start..].find(end_tag)?;
    Some(xml[start..start + end + end_tag.len()].to_string())
}

/// Generate body XML from content blocks (shared by build_docx and build_docx_with_template).
fn generate_body_xml(blocks: &[ContentBlock]) -> String {
    let mut body_xml = String::new();
    for block in blocks {
        match block.block_type.as_str() {
            "paragraph" => {
                let mut ppr = String::new();
                if let Some(level) = block.heading_level {
                    if level >= 1 && level <= 6 {
                        ppr.push_str(&format!(r#"<w:pStyle w:val="Heading{}"/>"#, level));
                    }
                }
                if let Some(ref align) = block.alignment {
                    let val = match align.as_str() {
                        "center" => "center",
                        "right" => "right",
                        "justify" => "both",
                        _ => "left",
                    };
                    ppr.push_str(&format!(r#"<w:jc w:val="{}"/>"#, val));
                }
                if let Some(lh) = block.line_height {
                    if (lh - 1.0).abs() > 0.01 {
                        let val = (lh * 240.0).round() as i32;
                        ppr.push_str(&format!(r#"<w:spacing w:line="{}" w:lineRule="auto"/>"#, val));
                    }
                }
                body_xml.push_str("<w:p>");
                if !ppr.is_empty() {
                    body_xml.push_str(&format!("<w:pPr>{}</w:pPr>", ppr));
                }
                if block.runs.is_empty() {
                    body_xml.push_str("<w:r><w:t></w:t></w:r>");
                } else {
                    for run in &block.runs {
                        let rpr = build_rpr(run);
                        let text = editor::escape_xml_public(&run.text);
                        let space = if run.text.starts_with(' ') || run.text.ends_with(' ') {
                            r#" xml:space="preserve""#
                        } else { "" };
                        body_xml.push_str(&format!("<w:r>{}<w:t{}>{}</w:t></w:r>", rpr, space, text));
                    }
                }
                body_xml.push_str("</w:p>");
            }
            "table" => {
                if let Some(ref rows) = block.rows {
                    let cols = rows.first().map(|r| r.len()).unwrap_or(1);
                    let col_w = 8640 / cols as i32;
                    body_xml.push_str("<w:tbl>");
                    body_xml.push_str(concat!(
                        "<w:tblPr>",
                        r#"<w:tblW w:type="auto" w:w="0"/>"#,
                        "<w:tblBorders>",
                        r#"<w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>"#,
                        r#"<w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>"#,
                        r#"<w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>"#,
                        r#"<w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>"#,
                        r#"<w:insideH w:val="single" w:sz="4" w:space="0" w:color="000000"/>"#,
                        r#"<w:insideV w:val="single" w:sz="4" w:space="0" w:color="000000"/>"#,
                        "</w:tblBorders>",
                        r#"<w:tblLook w:firstColumn="1" w:firstRow="1" w:lastColumn="0" w:lastRow="0" w:noHBand="0" w:noVBand="1" w:val="04A0"/>"#,
                        "</w:tblPr>"
                    ));
                    body_xml.push_str("<w:tblGrid>");
                    for _ in 0..cols {
                        body_xml.push_str(&format!(r#"<w:gridCol w:w="{}"/>"#, col_w));
                    }
                    body_xml.push_str("</w:tblGrid>");
                    for row in rows {
                        body_xml.push_str("<w:tr>");
                        for cell in row {
                            let text = editor::escape_xml_public(&cell.text);
                            let rpr = if cell.bold == Some(true) { "<w:rPr><w:b/><w:bCs/></w:rPr>" } else { "" };
                            body_xml.push_str(&format!(
                                r#"<w:tc><w:tcPr><w:tcW w:type="dxa" w:w="{}"/></w:tcPr><w:p><w:r>{}<w:t>{}</w:t></w:r></w:p></w:tc>"#,
                                col_w, rpr, text
                            ));
                        }
                        body_xml.push_str("</w:tr>");
                    }
                    body_xml.push_str("</w:tbl>");
                }
            }
            _ => {}
        }
    }
    if body_xml.is_empty() {
        body_xml = "<w:p><w:r><w:t></w:t></w:r></w:p>".to_string();
    }
    body_xml
}

fn build_rpr(run: &ContentRun) -> String {
    let mut parts = Vec::new();

    if let Some(ref ff) = run.font_family {
        parts.push(format!(r#"<w:rFonts w:ascii="{}" w:hAnsi="{}" w:eastAsia="{}"/>"#, ff, ff, ff));
    }
    if run.bold == Some(true) {
        parts.push("<w:b/><w:bCs/>".to_string());
    }
    if run.italic == Some(true) {
        parts.push("<w:i/><w:iCs/>".to_string());
    }
    if run.strikethrough == Some(true) {
        parts.push("<w:strike/>".to_string());
    }
    if let Some(ref color) = run.color {
        let c = color.trim_start_matches('#');
        parts.push(format!(r#"<w:color w:val="{}"/>"#, c));
    }
    if let Some(size) = run.font_size {
        let hp = (size * 2.0).round() as u32;
        parts.push(format!(r#"<w:sz w:val="{}"/><w:szCs w:val="{}"/>"#, hp, hp));
    }
    if run.underline == Some(true) {
        parts.push(r#"<w:u w:val="single"/>"#.to_string());
    }

    if parts.is_empty() {
        String::new()
    } else {
        format!("<w:rPr>{}</w:rPr>", parts.join(""))
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_default_page_size() {
        let size = ir::PageSize::default();
        assert!((size.width - 595.0).abs() < f32::EPSILON);
        assert!((size.height - 842.0).abs() < f32::EPSILON);
    }

    #[test]
    fn test_parse_basic_docx() {
        let data = include_bytes!("../../../tests/fixtures/basic_test.docx");
        let doc = parse_docx(data).expect("should parse basic docx");

        assert_eq!(doc.pages.len(), 1);
        let page = &doc.pages[0];

        // Check page size: A4 in twips (11906/20 = 595.3, 16838/20 = 841.9)
        assert!((page.size.width - 595.3).abs() < 0.1);
        assert!((page.size.height - 841.9).abs() < 0.1);

        // Check margins: 1440 twips = 72pt (1 inch)
        assert!((page.margin.top - 72.0).abs() < 0.1);
        assert!((page.margin.left - 72.0).abs() < 0.1);

        // Check blocks: heading + paragraph + japanese + table
        assert_eq!(page.blocks.len(), 4);

        // First block: heading paragraph
        match &page.blocks[0] {
            ir::Block::Paragraph(p) => {
                assert_eq!(p.runs.len(), 1);
                assert_eq!(p.runs[0].text, "Oxidocs Test Document");
                assert!(p.runs[0].style.bold);
                assert_eq!(p.style.heading_level, Some(1));
            }
            _ => panic!("expected paragraph"),
        }

        // Second block: mixed formatting paragraph
        match &page.blocks[1] {
            ir::Block::Paragraph(p) => {
                assert!(p.runs.len() >= 3);
                assert!(p.runs[1].style.bold); // "bold text"
                assert!(p.runs[3].style.italic); // "italic text"
            }
            _ => panic!("expected paragraph"),
        }

        // Third block: Japanese text
        match &page.blocks[2] {
            ir::Block::Paragraph(p) => {
                assert!(p.runs[0].text.contains("日本語"));
            }
            _ => panic!("expected paragraph"),
        }

        // Fourth block: table
        match &page.blocks[3] {
            ir::Block::Table(t) => {
                assert_eq!(t.rows.len(), 2);
                assert_eq!(t.rows[0].cells.len(), 2);
                assert!(t.style.border);
            }
            _ => panic!("expected table"),
        }
    }

    #[test]
    fn test_styles_parsed() {
        let data = include_bytes!("../../../tests/fixtures/basic_test.docx");
        let doc = parse_docx(data).expect("should parse");

        // Should have Normal and Heading1 styles
        assert!(doc.styles.styles.contains_key("Normal"));
        assert!(doc.styles.styles.contains_key("Heading1"));

        let h1 = &doc.styles.styles["Heading1"];
        assert_eq!(h1.paragraph.space_before, Some(12.0)); // 240 twips / 20 = 12pt
        assert_eq!(h1.paragraph.space_after, Some(6.0));    // 120 twips / 20 = 6pt
    }

    #[test]
    fn test_layout_basic() {
        let data = include_bytes!("../../../tests/fixtures/basic_test.docx");
        let doc = parse_docx(data).expect("should parse");
        let engine = layout::LayoutEngine::for_document(&doc);
        let result = engine.layout(&doc);

        // Should produce at least 1 page
        assert!(!result.pages.is_empty());
        let page = &result.pages[0];

        // Page should have elements
        assert!(!page.elements.is_empty());

        // Check page dimensions match A4
        assert!((page.width - 595.3).abs() < 0.1);
        assert!((page.height - 841.9).abs() < 0.1);

        // First element should be text starting at margin position (with heading space_before)
        let first = &page.elements[0];
        assert!((first.x - 72.0).abs() < 1.0); // left margin
        // y may be offset by heading space_before from style definition
        assert!(first.y >= 72.0 && first.y < 100.0); // top margin + possible heading spacing
    }

    #[test]
    fn test_create_blank_docx() {
        let bytes = create_blank_docx();
        assert!(!bytes.is_empty());

        // Should be parseable
        let doc = parse_docx(&bytes).expect("blank docx should parse");
        assert_eq!(doc.pages.len(), 1);

        // Should have at least one block
        let page = &doc.pages[0];
        assert!(!page.blocks.is_empty());

        // Should be A4 size
        assert!((page.size.width - 595.3).abs() < 0.1);
        assert!((page.size.height - 841.9).abs() < 0.1);

        // Should be editable
        let mut editor = DocxEditor::new(&bytes).expect("should create editor");
        editor.apply_edits(&[editor::TextEdit {
            paragraph_index: 0,
            run_index: 0,
            new_text: "Hello World".to_string(),
        }]);
        let saved = editor.save().expect("should save");
        let doc2 = parse_docx(&saved).expect("edited blank should parse");
        match &doc2.pages[0].blocks[0] {
            ir::Block::Paragraph(p) => {
                assert_eq!(p.runs[0].text, "Hello World");
            }
            _ => panic!("expected paragraph"),
        }
    }

    #[test]
    fn test_parse_docx_with_image() {
        let data = include_bytes!("../../../tests/fixtures/with_image.docx");
        let doc = parse_docx(data).expect("should parse docx with image");

        let page = &doc.pages[0];
        // 4 blocks: "Before image" paragraph, image paragraph, inline Image block, "After image" paragraph
        assert_eq!(page.blocks.len(), 4);

        // First paragraph: "Before image"
        match &page.blocks[0] {
            ir::Block::Paragraph(p) => {
                assert_eq!(p.runs[0].text, "Before image");
            }
            _ => panic!("expected paragraph"),
        }

        // Third block: inline image
        assert!(matches!(&page.blocks[2], ir::Block::Image(_)), "expected inline image block");

        // Fourth paragraph: "After image"
        match &page.blocks[3] {
            ir::Block::Paragraph(p) => {
                assert_eq!(p.runs[0].text, "After image");
            }
            _ => panic!("expected paragraph"),
        }
    }
}
