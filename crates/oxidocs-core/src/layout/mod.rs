use crate::ir::Document;

pub struct LayoutEngine;

impl LayoutEngine {
    pub fn new() -> Self {
        Self
    }

    pub fn layout(&self, _doc: &Document) -> LayoutResult {
        // TODO: Implement layout engine
        LayoutResult { pages: vec![] }
    }
}

pub struct LayoutResult {
    pub pages: Vec<LayoutPage>,
}

pub struct LayoutPage {
    pub elements: Vec<LayoutElement>,
}

pub struct LayoutElement {
    pub x: f32,
    pub y: f32,
    pub width: f32,
    pub height: f32,
    pub content: LayoutContent,
}

pub enum LayoutContent {
    Text { text: String, font_size: f32 },
    Image { data: Vec<u8> },
    Border { style: BorderStyle },
}

pub struct BorderStyle {
    pub width: f32,
}
