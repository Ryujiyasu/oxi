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
}
