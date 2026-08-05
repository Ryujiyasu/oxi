// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

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
