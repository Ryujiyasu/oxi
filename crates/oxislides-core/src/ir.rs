// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

use std::collections::HashMap;

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
    /// Theme colour scheme (a:clrScheme from ppt/theme/themeN.xml): scheme slot
    /// name (dk1/dk2/lt1/lt2/tx1/tx2/accent1..accent6/hlink/folHlink) -> RGB hex
    /// (6-digit RRGGBB). `<a:schemeClr val="..."/>` references resolve through
    /// this map first; the built-in Office table is only the fallback when the
    /// theme part is absent (Spec #10).
    #[serde(default)]
    pub theme_colors: HashMap<String, String>,
    /// Slide master text styles (p:txStyles from slideMaster1.xml): the
    /// inherited marL/indent/bullet/spcBef per outline level for body / other
    /// (textbox) / title contexts (Spec #8). Placeholder body text uses
    /// `body`, plain textboxes use `other`, title placeholders use `title`.
    #[serde(default)]
    pub master_styles: MasterTxStyles,
}

/// Slide master text styles (p:txStyles). Each Vec is indexed by outline level
/// (0-based, a:lvlNpPr where N = level+1).
#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct MasterTxStyles {
    #[serde(default)]
    pub body: Vec<MasterStyleLevel>,
    #[serde(default)]
    pub other: Vec<MasterStyleLevel>,
    #[serde(default)]
    pub title: Vec<MasterStyleLevel>,
}

/// One outline-level master text style (a:lvlNpPr).
#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct MasterStyleLevel {
    /// a:lvlNpPr/@marL in points (left indent of the paragraph; default 0).
    #[serde(default)]
    pub mar_l: f32,
    /// a:lvlNpPr/@indent in points (first-line indent; negative = hanging).
    #[serde(default)]
    pub indent: f32,
    /// Inherited bullet marker (default Inherit = none unless a bu* child).
    #[serde(default)]
    pub bullet: SlideBullet,
    /// a:spcBef/a:spcPct val/100000 — space-before as a fraction of the line
    /// advance (e.g. 0.2 = 20000). None = no spcBef on this level.
    #[serde(default)]
    pub spc_bef_pct: Option<f32>,
    /// a:defRPr/@sz in points — the placeholder default font size for this
    /// outline level. None = no explicit size (engine default 18pt applies).
    /// Word render-truth (phfs probe, 2026-08): a body placeholder inherits
    /// the MASTER txStyles bodyStyle level size (layout txStyles is ignored);
    /// a title placeholder inherits master titleStyle. A run's explicit sz
    /// always wins.
    #[serde(default)]
    pub font_size: Option<f32>,
    /// a:lvlNpPr/@algn — the horizontal paragraph alignment inherited from the
    /// master txStyles level (Spec #6). None = not specified (paragraph-level
    /// Left applies). The master titleStyle lvl1pPr carries algn="ctr", which
    /// is what horizontally centres title placeholders (V4/P2 render-truth).
    #[serde(default)]
    pub algn: Option<SlideAlignment>,
}

pub fn default_l_ins() -> f32 {
    7.2
}
pub fn default_r_ins() -> f32 {
    7.2
}
pub fn default_t_ins() -> f32 {
    3.6
}
pub fn default_b_ins() -> f32 {
    3.6
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
    /// Text-area insets from a:bodyPr (lIns/rIns/tIns/bIns) in points.
    /// A placeholder with no bodyPr inset uses 7.2 / 7.2 / 3.6 / 3.6; a textbox
    /// carries its own insets (e.g. lIns=914400 EMU = 72pt). The left text
    /// position P0 = shape.x + l_ins (Spec #8).
    #[serde(default = "default_l_ins")]
    pub l_ins: f32,
    #[serde(default = "default_r_ins")]
    pub r_ins: f32,
    #[serde(default = "default_t_ins")]
    pub t_ins: f32,
    #[serde(default = "default_b_ins")]
    pub b_ins: f32,
    /// Vertical text-anchor (a:bodyPr/@anchor), resolved through the
    /// placeholder chain (slide shape's own bodyPr -> layout placeholder ->
    /// master placeholder). None = "t" (top). Spec #6 render-truth: the master
    /// TITLE placeholder carries anchor="ctr", which vertically centres title
    /// placeholders; layout3 title uses "t", layout8/9 "b".
    #[serde(default)]
    pub anchor: Option<String>,
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
    /// Horizontal alignment (a:pPr/@algn). None = not specified on the
    /// paragraph — the master txStyles level alignment (MasterStyleLevel.algn)
    /// applies at render time (Spec #6).
    #[serde(default)]
    pub alignment: Option<SlideAlignment>,
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
    /// Outline level (a:pPr/@lvl, 0-based). Default 0.
    #[serde(default)]
    pub lvl: u32,
    /// Left indent (a:pPr/@marL) in points. None = not specified (the master
    /// txStyles level provides it).
    #[serde(default)]
    pub mar_l: Option<f32>,
    /// First-line indent (a:pPr/@indent) in points. None = not specified.
    #[serde(default)]
    pub indent: Option<f32>,
    /// Bullet marker spec (a:buChar / a:buNone / a:buAutoNum). Inherit = use
    /// the master txStyles level's bullet.
    #[serde(default)]
    pub bullet: SlideBullet,
}

/// Bullet marker specification for a paragraph (Spec #8, measured model).
#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub enum SlideBullet {
    /// No a:bu* child on the paragraph itself: inherit from the style chain
    /// (the master txStyles level for the shape's context).
    #[default]
    Inherit,
    /// a:buNone — no bullet marker is drawn (but indent geometry still applies).
    None,
    /// a:buChar — a literal character marker (e.g. "•", "–", "»").
    Char {
        ch: char,
        font: Option<String>, // a:buFont/@typeface
    },
    /// a:buAutoNum — an automatically numbered marker (rendering: follow-up).
    AutoNum {
        kind: String,
    },
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
