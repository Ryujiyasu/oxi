use serde::{Deserialize, Serialize};
use std::collections::HashMap;

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct FontMetrics {
    pub family: String,
    pub size: f32,
    pub ascent: f32,
    pub descent: f32,
    pub line_gap: f32,
    pub char_widths: HashMap<char, f32>,
}

impl FontMetrics {
    pub fn char_width(&self, c: char) -> f32 {
        self.char_widths.get(&c).copied().unwrap_or(self.size * 0.5)
    }

    pub fn line_height(&self) -> f32 {
        self.ascent + self.descent + self.line_gap
    }
}
