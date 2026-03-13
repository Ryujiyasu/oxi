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
        let engine = layout::LayoutEngine::new();
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
