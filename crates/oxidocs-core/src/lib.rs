pub mod font;
pub mod ir;
pub mod layout;
pub mod parser;

pub use ir::Document;
pub use parser::parse_docx;

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
        assert_eq!(h1.space_before, Some(12.0)); // 240 twips / 20 = 12pt
        assert_eq!(h1.space_after, Some(6.0));    // 120 twips / 20 = 6pt
    }
}
