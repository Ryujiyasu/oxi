pub mod ir;
pub mod parser;

pub use parser::parse_pptx;

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_parse_basic_pptx() {
        let data = include_bytes!("../../../tests/fixtures/basic_test.pptx");
        let pres = parse_pptx(data).expect("parse failed");
        assert_eq!(pres.slides.len(), 1);
        assert!((pres.slide_width - 720.0).abs() < 1.0);
        assert!((pres.slide_height - 540.0).abs() < 1.0);

        let slide = &pres.slides[0];
        assert!(slide.shapes.len() >= 2); // title + body

        // Check title shape
        let title = &slide.shapes[0];
        if let ir::ShapeContent::TextBox { paragraphs } = &title.content {
            assert!(!paragraphs.is_empty());
            let text: String = paragraphs[0].runs.iter().map(|r| r.text.as_str()).collect();
            assert_eq!(text, "Welcome to Oxi");
            assert!(paragraphs[0].runs[0].bold);
        } else {
            panic!("Expected TextBox for title");
        }

        // Check body has multiple paragraphs
        let body = &slide.shapes[1];
        if let ir::ShapeContent::TextBox { paragraphs } = &body.content {
            assert!(paragraphs.len() >= 3);
            // Check italic + colored run
            let rust_run = &paragraphs[1].runs[0];
            assert!(rust_run.italic);
            assert_eq!(rust_run.color.as_deref(), Some("4472C4"));
        } else {
            panic!("Expected TextBox for body");
        }
    }

    #[test]
    fn test_parse_multi_slide_pptx() {
        let data = include_bytes!("../../../tests/fixtures/multi_slide.pptx");
        let pres = parse_pptx(data).expect("parse failed");
        assert_eq!(pres.slides.len(), 3);

        // Slide 1 — title
        let s1 = &pres.slides[0];
        if let ir::ShapeContent::TextBox { paragraphs } = &s1.shapes[0].content {
            let text: String = paragraphs[0].runs.iter().map(|r| r.text.as_str()).collect();
            assert_eq!(text, "Oxi Project");
        } else {
            panic!("Expected TextBox");
        }

        // Slide 2 — bullet points
        let s2 = &pres.slides[1];
        assert!(s2.shapes.len() >= 2);

        // Slide 3 — Japanese text
        let s3 = &pres.slides[2];
        if let ir::ShapeContent::TextBox { paragraphs } = &s3.shapes[1].content {
            assert!(paragraphs.len() >= 3);
            // Check font family
            let first_run = &paragraphs[0].runs[0];
            assert!(first_run.text.contains('猫') || first_run.text.contains('吾'));
        } else {
            panic!("Expected TextBox");
        }
    }
}
