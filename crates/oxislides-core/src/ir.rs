// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

use serde::{Deserialize, Serialize};

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Presentation {
    pub slides: Vec<Slide>,
    pub slide_width: f32,  // in points (default 960pt = 10 inches)
    pub slide_height: f32, // in points (default 540pt = 7.5 inches)
    /// Theme minor-font latin typeface (default font for body text / most shapes).
    /// From ppt/theme/themeN.xml <a:minorFont><a:latin typeface="..."/>. Default "Calibri".
    #[serde(default = "default_theme_font")]
    pub minor_font: String,
    /// Theme major-font latin typeface (title placeholders). Default "Calibri".
    #[serde(default = "default_theme_font")]
    pub major_font: String,
}

pub fn default_theme_font() -> String {
    "Calibri".to_string()
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Slide {
    pub index: usize,
    pub shapes: Vec<Shape>,
    pub background_color: Option<String>, // hex color e.g. "FFFFFF"
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Shape {
    pub x: f32,      // position in points
    pub y: f32,
    pub width: f32,
    pub height: f32,
    pub rotation: f32, // rotation in degrees (0 = none), from a:xfrm/@rot (60000 = 1 deg)
    /// PresentationML shape type (a:prstGeom/@prst e.g. "rect", "ellipse", "roundRect", "chevron").
    /// None for picture / graphicFrame / plain textbox without a preset geometry.
    pub shape_type: Option<String>,
    /// Placeholder type (p:nvPr/p:ph/@type, "title" / "body" / "subTitle" / "obj"...).
    /// None for a plain (non-placeholder) shape. Title placeholders use the theme
    /// MAJOR font; everything else (incl. body placeholders) uses the MINOR font.
    #[serde(default)]
    pub ph_type: Option<String>,
    pub content: ShapeContent,
    pub fill_color: Option<String>,   // hex color for solid fill
    pub border_color: Option<String>, // hex color for outline
    pub border_width: Option<f32>,    // border width in points
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub enum ShapeContent {
    /// A preset-geometry AutoShape (a:prstGeom). May carry text (a:txBody).
    AutoShape {
        paragraphs: Vec<SlideParagraph>,
    },
    TextBox {
        paragraphs: Vec<SlideParagraph>,
    },
    /// A DrawingML table (a:graphicFrame -> a:graphicData uri=.../table -> a:tbl).
    Table {
        table: Table,
    },
    Image {
        data: Vec<u8>,
        content_type: Option<String>,
    },
    /// Unsupported element with type label (e.g. "SmartArt", "Chart", "OLE")
    Unsupported {
        element_type: String,
    },
    Placeholder, // shapes we can't parse yet
}

/// A DrawingML table (a:tbl). Cell text is stored as paragraphs per cell so
/// the run/paragraph machinery is shared with shapes.
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Table {
    /// Column widths in points (a:tblGrid/a:gridCol/@w).
    pub col_widths: Vec<f32>,
    /// Row heights in points (a:tr/@h).
    pub row_heights: Vec<f32>,
    /// Cells in row-major order: rows[r][c].
    pub rows: Vec<Vec<TableCell>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct TableCell {
    pub paragraphs: Vec<SlideParagraph>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SlideParagraph {
    pub runs: Vec<SlideRun>,
    pub alignment: SlideAlignment,
    /// Multiple line-spacing factor `n` from `a:lnSpc/a:spcPct` (val/100000).
    /// None = no explicit lnSpc -> PowerPoint uses its default single line
    /// (measured line advance = fs*1.2).  When Some(n), the measured line
    /// advance = fs*1.2*n (linear over n in [0.5, 3.0], verified spec4d).
    #[serde(default)]
    pub line_spacing: Option<f32>,
    /// Space before this paragraph, in points (`a:spcBef/a:spcPts` val/100).
    /// Added on top of the line advance (wave-1 measurement).
    #[serde(default)]
    pub space_before: Option<f32>,
    /// Space after this paragraph, in points (`a:spcAft/a:spcPts` val/100).
    #[serde(default)]
    pub space_after: Option<f32>,
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
