//! Hanko stamp image generation.
//!
//! Generates traditional Japanese seal (hanko/inkan) stamp images as SVG.
//! Supports round (丸印) and square (角印) styles.

use serde::{Deserialize, Serialize};

/// Style of the hanko stamp.
#[derive(Debug, Clone, Copy, PartialEq, Serialize, Deserialize)]
pub enum StampStyle {
    /// Round stamp (丸印) — most common for personal use.
    Round,
    /// Square stamp (角印) — used for company seals.
    Square,
    /// Oval stamp (小判型) — used for bank registration.
    Oval,
}

/// Color of the stamp ink.
#[derive(Debug, Clone, Copy, Serialize, Deserialize)]
pub struct StampColor {
    pub r: u8,
    pub g: u8,
    pub b: u8,
}

impl StampColor {
    /// Traditional vermilion red (朱色).
    pub fn vermilion() -> Self {
        Self { r: 227, g: 66, b: 52 }
    }

    /// Standard red.
    pub fn red() -> Self {
        Self { r: 220, g: 30, b: 30 }
    }

    /// Black ink.
    pub fn black() -> Self {
        Self { r: 0, g: 0, b: 0 }
    }

    fn to_hex(&self) -> String {
        format!("#{:02X}{:02X}{:02X}", self.r, self.g, self.b)
    }
}

/// Configuration for generating a hanko stamp.
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct StampConfig {
    /// Name to display on the stamp (typically 1-4 kanji characters).
    pub name: String,
    /// Stamp style.
    pub style: StampStyle,
    /// Stamp color.
    pub color: StampColor,
    /// Size in pixels (diameter for round, side length for square).
    pub size: u32,
    /// Border width as a fraction of the size (0.0 - 1.0).
    pub border_width: f64,
    /// Font size as a fraction of the size (0.0 - 1.0).
    pub font_ratio: f64,
    /// Date string to display below the name (optional, for approval stamps).
    pub date: Option<String>,
}

impl Default for StampConfig {
    fn default() -> Self {
        Self {
            name: String::new(),
            style: StampStyle::Round,
            color: StampColor::vermilion(),
            size: 100,
            border_width: 0.08,
            font_ratio: 0.45,
            date: None,
        }
    }
}

/// Generate a hanko stamp as an SVG string.
pub fn generate_stamp_svg(config: &StampConfig) -> String {
    match config.style {
        StampStyle::Round => generate_round_stamp(config),
        StampStyle::Square => generate_square_stamp(config),
        StampStyle::Oval => generate_oval_stamp(config),
    }
}

fn generate_round_stamp(config: &StampConfig) -> String {
    let size = config.size;
    let cx = size as f64 / 2.0;
    let cy = size as f64 / 2.0;
    let radius = cx - 2.0;
    let border = size as f64 * config.border_width;
    let color = config.color.to_hex();

    let font_size = size as f64 * config.font_ratio;
    let chars: Vec<char> = config.name.chars().collect();

    let text_elements = if let Some(ref date) = config.date {
        // Approval stamp: name on top, date in middle, department on bottom.
        let name_y = cy - font_size * 0.3;
        let date_font_size = font_size * 0.5;
        let date_y = cy + date_font_size * 0.2;
        format!(
            r#"    <line x1="{x1}" y1="{ly1}" x2="{x2}" y2="{ly1}" stroke="{color}" stroke-width="{sw}"/>
    <line x1="{x1}" y1="{ly2}" x2="{x2}" y2="{ly2}" stroke="{color}" stroke-width="{sw}"/>
    <text x="{cx}" y="{name_y}" text-anchor="middle" dominant-baseline="middle"
          fill="{color}" font-family="serif" font-size="{font_size}">{name}</text>
    <text x="{cx}" y="{date_y}" text-anchor="middle" dominant-baseline="middle"
          fill="{color}" font-family="serif" font-size="{date_font_size}">{date}</text>"#,
            x1 = cx - radius * 0.7,
            x2 = cx + radius * 0.7,
            ly1 = cy - font_size * 0.55,
            ly2 = cy + font_size * 0.15,
            sw = border * 0.5,
            name = config.name,
            date = date,
        )
    } else if chars.len() <= 2 {
        // Short name: horizontal layout.
        format!(
            r#"    <text x="{cx}" y="{cy}" text-anchor="middle" dominant-baseline="central"
          fill="{color}" font-family="serif" font-size="{font_size}"
          letter-spacing="{ls}">{name}</text>"#,
            ls = font_size * 0.15,
            name = config.name,
        )
    } else {
        // Longer name: vertical layout.
        let mut elements = String::new();
        let total_height = chars.len() as f64 * font_size * 0.9;
        let start_y = cy - total_height / 2.0 + font_size * 0.45;
        for (i, ch) in chars.iter().enumerate() {
            let y = start_y + i as f64 * font_size * 0.9;
            elements.push_str(&format!(
                r#"    <text x="{cx}" y="{y}" text-anchor="middle" dominant-baseline="central"
          fill="{color}" font-family="serif" font-size="{font_size}">{ch}</text>
"#,
            ));
        }
        elements
    };

    format!(
        r#"<svg xmlns="http://www.w3.org/2000/svg" width="{size}" height="{size}" viewBox="0 0 {size} {size}">
    <circle cx="{cx}" cy="{cy}" r="{radius}" fill="none" stroke="{color}" stroke-width="{border}"/>
{text_elements}
</svg>"#,
    )
}

fn generate_square_stamp(config: &StampConfig) -> String {
    let size = config.size;
    let border = size as f64 * config.border_width;
    let color = config.color.to_hex();
    let font_size = size as f64 * config.font_ratio;
    let cx = size as f64 / 2.0;
    let cy = size as f64 / 2.0;

    let chars: Vec<char> = config.name.chars().collect();

    // For square stamps, arrange characters in a grid.
    let text_elements = if chars.len() <= 4 {
        // 2x2 grid layout (common for company seals).
        let cols = 2;
        let rows = (chars.len() + 1) / 2;
        let cell_w = (size as f64 - border * 4.0) / cols as f64;
        let cell_h = (size as f64 - border * 4.0) / rows as f64;
        let char_size = cell_w.min(cell_h) * 0.8;

        let mut elements = String::new();
        for (i, ch) in chars.iter().enumerate() {
            // Traditional right-to-left, top-to-bottom order.
            let col = 1 - (i % cols); // right to left
            let row = i / cols;
            let x = border * 2.0 + col as f64 * cell_w + cell_w / 2.0;
            let y = border * 2.0 + row as f64 * cell_h + cell_h / 2.0;
            elements.push_str(&format!(
                r#"    <text x="{x}" y="{y}" text-anchor="middle" dominant-baseline="central"
          fill="{color}" font-family="serif" font-size="{char_size}">{ch}</text>
"#,
            ));
        }
        elements
    } else {
        format!(
            r#"    <text x="{cx}" y="{cy}" text-anchor="middle" dominant-baseline="central"
          fill="{color}" font-family="serif" font-size="{font_size}">{name}</text>"#,
            name = config.name,
        )
    };

    format!(
        r#"<svg xmlns="http://www.w3.org/2000/svg" width="{size}" height="{size}" viewBox="0 0 {size} {size}">
    <rect x="{b}" y="{b}" width="{w}" height="{w}" fill="none" stroke="{color}" stroke-width="{border}"/>
{text_elements}
</svg>"#,
        b = border / 2.0,
        w = size as f64 - border,
    )
}

fn generate_oval_stamp(config: &StampConfig) -> String {
    let width = config.size;
    let height = (config.size as f64 * 0.7) as u32;
    let cx = width as f64 / 2.0;
    let cy = height as f64 / 2.0;
    let rx = cx - 2.0;
    let ry = cy - 2.0;
    let border = config.size as f64 * config.border_width;
    let color = config.color.to_hex();
    let font_size = height as f64 * config.font_ratio;

    format!(
        r#"<svg xmlns="http://www.w3.org/2000/svg" width="{width}" height="{height}" viewBox="0 0 {width} {height}">
    <ellipse cx="{cx}" cy="{cy}" rx="{rx}" ry="{ry}" fill="none" stroke="{color}" stroke-width="{border}"/>
    <text x="{cx}" y="{cy}" text-anchor="middle" dominant-baseline="central"
          fill="{color}" font-family="serif" font-size="{font_size}">{name}</text>
</svg>"#,
        name = config.name,
    )
}

/// Convert an SVG stamp to a simple bitmap representation.
/// Returns raw RGBA pixel data (for embedding as PDF image).
/// This is a placeholder — full SVG rasterization will use a proper renderer.
pub fn stamp_to_rgba(_svg: &str, size: u32) -> Vec<u8> {
    // Placeholder: return a solid red circle on transparent background.
    let mut pixels = vec![0u8; (size * size * 4) as usize];
    let cx = size as f64 / 2.0;
    let cy = size as f64 / 2.0;
    let r = cx - 2.0;

    for y in 0..size {
        for x in 0..size {
            let dx = x as f64 - cx;
            let dy = y as f64 - cy;
            let dist = (dx * dx + dy * dy).sqrt();

            let idx = ((y * size + x) * 4) as usize;
            if dist <= r && dist >= r - r * 0.08 {
                // Border ring
                pixels[idx] = 227; // R
                pixels[idx + 1] = 66; // G
                pixels[idx + 2] = 52; // B
                pixels[idx + 3] = 255; // A
            }
        }
    }

    pixels
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_round_stamp_svg() {
        let config = StampConfig {
            name: "山田".into(),
            style: StampStyle::Round,
            ..Default::default()
        };
        let svg = generate_stamp_svg(&config);
        assert!(svg.contains("<svg"));
        assert!(svg.contains("<circle"));
        assert!(svg.contains("山田"));
        assert!(svg.contains("#E34234")); // vermilion
    }

    #[test]
    fn test_square_stamp_svg() {
        let config = StampConfig {
            name: "株式会社".into(),
            style: StampStyle::Square,
            ..Default::default()
        };
        let svg = generate_stamp_svg(&config);
        assert!(svg.contains("<svg"));
        assert!(svg.contains("<rect"));
        assert!(svg.contains("株"));
        assert!(svg.contains("式"));
    }

    #[test]
    fn test_oval_stamp_svg() {
        let config = StampConfig {
            name: "鈴木".into(),
            style: StampStyle::Oval,
            ..Default::default()
        };
        let svg = generate_stamp_svg(&config);
        assert!(svg.contains("<svg"));
        assert!(svg.contains("<ellipse"));
        assert!(svg.contains("鈴木"));
    }

    #[test]
    fn test_approval_stamp() {
        let config = StampConfig {
            name: "承認".into(),
            style: StampStyle::Round,
            date: Some("2026.03.13".into()),
            ..Default::default()
        };
        let svg = generate_stamp_svg(&config);
        assert!(svg.contains("承認"));
        assert!(svg.contains("2026.03.13"));
        assert!(svg.contains("<line")); // divider lines
    }

    #[test]
    fn test_vertical_name() {
        let config = StampConfig {
            name: "田中太郎".into(),
            style: StampStyle::Round,
            ..Default::default()
        };
        let svg = generate_stamp_svg(&config);
        assert!(svg.contains("田"));
        assert!(svg.contains("中"));
        assert!(svg.contains("太"));
        assert!(svg.contains("郎"));
    }

    #[test]
    fn test_stamp_to_rgba() {
        let pixels = stamp_to_rgba("", 10);
        assert_eq!(pixels.len(), 10 * 10 * 4);
    }

    #[test]
    fn test_custom_color() {
        let config = StampConfig {
            name: "佐藤".into(),
            color: StampColor::black(),
            ..Default::default()
        };
        let svg = generate_stamp_svg(&config);
        assert!(svg.contains("#000000"));
    }
}
