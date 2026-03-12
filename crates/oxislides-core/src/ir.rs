use serde::{Deserialize, Serialize};

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Presentation {
    pub slides: Vec<Slide>,
    pub slide_width: f32,  // in points (default 960pt = 10 inches)
    pub slide_height: f32, // in points (default 540pt = 7.5 inches)
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Slide {
    pub index: usize,
    pub shapes: Vec<Shape>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Shape {
    pub x: f32,      // position in points
    pub y: f32,
    pub width: f32,
    pub height: f32,
    pub content: ShapeContent,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub enum ShapeContent {
    TextBox {
        paragraphs: Vec<SlideParagraph>,
    },
    Image {
        data: Vec<u8>,
        content_type: Option<String>,
    },
    Placeholder, // shapes we can't parse yet
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SlideParagraph {
    pub runs: Vec<SlideRun>,
    pub alignment: SlideAlignment,
}

#[derive(Debug, Clone, Copy, Default, Serialize, Deserialize)]
pub enum SlideAlignment {
    #[default]
    Left,
    Center,
    Right,
    Justify,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SlideRun {
    pub text: String,
    pub font_size: Option<f32>,    // in points
    pub bold: bool,
    pub italic: bool,
    pub color: Option<String>,     // hex color
    pub font_family: Option<String>,
}
