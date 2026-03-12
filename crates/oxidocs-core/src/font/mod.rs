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
    /// Approximate metrics for a Calibri-like Latin font at 11pt.
    /// These will be replaced by real measured metrics later.
    pub fn default_latin() -> Self {
        let size = 11.0;
        let mut char_widths = HashMap::new();

        // Approximate Calibri character widths at 11pt (in points)
        // Based on typical proportional font metrics
        for ch in ' '..='~' {
            let width = match ch {
                ' ' => 0.25,
                'i' | 'l' | '!' | '|' | '.' | ',' | ':' | ';' | '\'' => 0.3,
                'f' | 'j' | 't' | 'r' => 0.35,
                'm' | 'w' | 'M' | 'W' => 0.7,
                'A'..='Z' => 0.6,
                _ => 0.5,
            };
            char_widths.insert(ch, size * width);
        }

        // CJK characters are typically full-width (equal to font size)
        // We don't enumerate them all; char_width() falls back for unknowns

        Self {
            family: "Calibri".to_string(),
            size,
            ascent: size * 0.75,
            descent: size * 0.25,
            line_gap: size * 0.1,
            char_widths,
        }
    }

    pub fn char_width(&self, c: char) -> f32 {
        self.char_widths.get(&c).copied().unwrap_or_else(|| {
            // CJK characters are full-width
            if is_fullwidth(c) {
                self.size
            } else {
                self.size * 0.5
            }
        })
    }

    pub fn line_height(&self) -> f32 {
        self.ascent + self.descent + self.line_gap
    }
}

fn is_fullwidth(ch: char) -> bool {
    matches!(ch as u32,
        0x3000..=0x303F |  // CJK Symbols and Punctuation
        0x3040..=0x309F |  // Hiragana
        0x30A0..=0x30FF |  // Katakana
        0x4E00..=0x9FFF |  // CJK Unified Ideographs
        0xFF00..=0xFFEF    // Fullwidth Forms
    )
}
