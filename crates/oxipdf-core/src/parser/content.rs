//! PDF content stream parser.
//!
//! A content stream is a sequence of operators with their operands that
//! describe text, graphics, and images on a page.

use std::collections::HashMap;

use crate::error::PdfError;
use crate::ir::*;
use super::object::PdfObject;
use super::cmap::CMap;
use super::encoding::{FontEncoding, decode_with_encoding};

/// Pre-resolved page resources needed for content stream interpretation.
#[derive(Default)]
pub struct PageResources {
    /// Font name → ToUnicode CMap.
    pub font_cmaps: HashMap<String, CMap>,
    /// Font name → encoding (when no CMap is available).
    pub font_encodings: HashMap<String, FontEncoding>,
    /// Font resource key (e.g. "F1") → BaseFont name (e.g. "Helvetica-Bold").
    pub font_base_names: HashMap<String, String>,
    /// XObject name → (decoded stream bytes, matrix, optional sub-resources).
    /// Form XObjects contain their own content streams.
    pub xobject_streams: HashMap<String, XObjectData>,
}

/// Data for a resolved Form XObject.
pub struct XObjectData {
    /// Decoded content stream bytes.
    pub stream: Vec<u8>,
    /// Transformation matrix from /Matrix entry (default: identity).
    pub matrix: [f64; 6],
    /// Optional BBox for clipping.
    pub bbox: Option<Rectangle>,
    /// Resources inherited or defined in the XObject.
    pub resources: Option<Box<PageResources>>,
}

/// A parsed content-stream operator with its operands.
#[derive(Debug, Clone)]
pub struct Operator {
    pub name: String,
    pub operands: Vec<PdfObject>,
}

/// Graphics state tracked while interpreting a content stream.
#[derive(Debug, Clone)]
struct GraphicsState {
    // Text state
    font_name: String,
    font_size: f64,
    text_matrix: [f64; 6],
    line_matrix: [f64; 6],
    text_leading: f64,
    char_spacing: f64,
    word_spacing: f64,
    text_rise: f64,
    horiz_scaling: f64,
    // Color state
    fill_color: Color,
    stroke_color: Color,
    // Line/path state
    line_width: f64,
    line_cap: LineCap,
    line_join: LineJoin,
    miter_limit: f64,
    // CTM (current transformation matrix)
    ctm: [f64; 6],
}

impl Default for GraphicsState {
    fn default() -> Self {
        Self {
            font_name: String::new(),
            font_size: 12.0,
            text_matrix: [1.0, 0.0, 0.0, 1.0, 0.0, 0.0],
            line_matrix: [1.0, 0.0, 0.0, 1.0, 0.0, 0.0],
            text_leading: 0.0,
            char_spacing: 0.0,
            word_spacing: 0.0,
            text_rise: 0.0,
            horiz_scaling: 100.0,
            fill_color: Color::Gray(0.0), // black
            stroke_color: Color::Gray(0.0),
            line_width: 1.0,
            line_cap: LineCap::Butt,
            line_join: LineJoin::Miter,
            miter_limit: 10.0,
            ctm: [1.0, 0.0, 0.0, 1.0, 0.0, 0.0],
        }
    }
}

/// Interpret a content stream and extract content elements.
/// Coordinates are in raw PDF space (origin at bottom-left).
pub fn interpret_content_stream(
    stream_data: &[u8],
    _page_height: f64,
) -> Result<Vec<ContentElement>, PdfError> {
    let resources = PageResources::default();
    interpret_content_stream_impl(stream_data, &resources, 0)
}

/// Interpret a content stream with full page resources (CMaps, encodings, XObjects).
pub fn interpret_content_stream_with_resources(
    stream_data: &[u8],
    resources: &PageResources,
) -> Result<Vec<ContentElement>, PdfError> {
    interpret_content_stream_impl(stream_data, resources, 0)
}

/// Interpret a content stream with font-specific CMap decoders (legacy API).
#[allow(dead_code)]
pub fn interpret_content_stream_with_cmaps(
    stream_data: &[u8],
    _page_height: f64,
    font_cmaps: &HashMap<String, CMap>,
) -> Result<Vec<ContentElement>, PdfError> {
    let resources = PageResources {
        font_cmaps: font_cmaps.clone(),
        ..Default::default()
    };
    interpret_content_stream_impl(stream_data, &resources, 0)
}

/// Maximum recursion depth for Form XObject interpretation.
const MAX_XOBJECT_DEPTH: u32 = 10;

fn interpret_content_stream_impl(
    stream_data: &[u8],
    resources: &PageResources,
    depth: u32,
) -> Result<Vec<ContentElement>, PdfError> {
    if depth > MAX_XOBJECT_DEPTH {
        return Ok(Vec::new());
    }
    let operators = tokenize_content_stream(stream_data)?;
    let mut elements = Vec::new();
    let mut state = GraphicsState::default();
    let mut state_stack: Vec<GraphicsState> = Vec::new();
    let mut current_path: Vec<PathOp> = Vec::new();
    let mut pending_clip: Option<bool> = None; // Some(even_odd)

    // Helper: build StrokeStyle from current state
    let make_stroke = |state: &GraphicsState| -> StrokeStyle {
        StrokeStyle {
            color: state.stroke_color,
            width: state.line_width,
            line_cap: state.line_cap,
            line_join: state.line_join,
        }
    };

    // Helper: compute effective font size (accounts for text matrix + CTM scaling)
    let effective_font_size = |state: &GraphicsState| -> f64 {
        let tm = &state.text_matrix;
        let ctm = &state.ctm;
        let combined = multiply_matrix(tm, ctm);
        // Scale factor from the y-axis of the combined matrix
        let sy = (combined[2] * combined[2] + combined[3] * combined[3]).sqrt();
        (state.font_size * sy).abs()
    };

    // Helper: emit a text element
    let emit_text = |state: &GraphicsState, text: String, elements: &mut Vec<ContentElement>| {
        if !text.trim().is_empty() {
            let (x, y) = transform_point(&state.ctm, &state.text_matrix);
            let font_size = effective_font_size(state);
            // Use BaseFont name if available, otherwise keep resource key
            let font_name = resources
                .font_base_names
                .get(&state.font_name)
                .cloned()
                .unwrap_or_else(|| state.font_name.clone());
            elements.push(ContentElement::Text(TextSpan {
                x,
                y,
                text,
                font_name,
                font_size,
                fill_color: state.fill_color,
            }));
        }
    };

    for op in &operators {
        match op.name.as_str() {
            // --- Graphics state ---
            "q" => {
                state_stack.push(state.clone());
                elements.push(ContentElement::SaveState);
            }
            "Q" => {
                if let Some(saved) = state_stack.pop() {
                    state = saved;
                }
                elements.push(ContentElement::RestoreState);
            }
            "cm" => {
                if op.operands.len() >= 6 {
                    let m = extract_matrix(&op.operands);
                    state.ctm = multiply_matrix(&state.ctm, &m);
                }
            }
            // Line width
            "w" => {
                if let Some(w) = op.operands.first().and_then(|o| o.as_f64()) {
                    state.line_width = w;
                }
            }
            // Line cap style: 0=Butt, 1=Round, 2=Square
            "J" => {
                if let Some(v) = op.operands.first().and_then(|o| o.as_i64()) {
                    state.line_cap = match v {
                        1 => LineCap::Round,
                        2 => LineCap::Square,
                        _ => LineCap::Butt,
                    };
                }
            }
            // Line join style: 0=Miter, 1=Round, 2=Bevel
            "j" => {
                if let Some(v) = op.operands.first().and_then(|o| o.as_i64()) {
                    state.line_join = match v {
                        1 => LineJoin::Round,
                        2 => LineJoin::Bevel,
                        _ => LineJoin::Miter,
                    };
                }
            }
            // Miter limit
            "M" => {
                if let Some(v) = op.operands.first().and_then(|o| o.as_f64()) {
                    state.miter_limit = v;
                }
            }
            // Dash pattern (ignored for now, but don't fail)
            "d" | "i" | "gs" | "ri" => {}

            // --- Text state operators ---
            "BT" => {
                state.text_matrix = [1.0, 0.0, 0.0, 1.0, 0.0, 0.0];
                state.line_matrix = [1.0, 0.0, 0.0, 1.0, 0.0, 0.0];
            }
            "ET" => {}
            "Tf" => {
                if op.operands.len() >= 2 {
                    if let Some(name) = op.operands[0].as_name() {
                        state.font_name = name.to_string();
                    }
                    if let Some(size) = op.operands[1].as_f64() {
                        state.font_size = size;
                    }
                }
            }
            // Text leading
            "TL" => {
                if let Some(leading) = op.operands.first().and_then(|o| o.as_f64()) {
                    state.text_leading = leading;
                }
            }
            // Character spacing
            "Tc" => {
                if let Some(v) = op.operands.first().and_then(|o| o.as_f64()) {
                    state.char_spacing = v;
                }
            }
            // Word spacing
            "Tw" => {
                if let Some(v) = op.operands.first().and_then(|o| o.as_f64()) {
                    state.word_spacing = v;
                }
            }
            // Horizontal scaling
            "Tz" => {
                if let Some(v) = op.operands.first().and_then(|o| o.as_f64()) {
                    state.horiz_scaling = v;
                }
            }
            // Text rise
            "Ts" => {
                if let Some(v) = op.operands.first().and_then(|o| o.as_f64()) {
                    state.text_rise = v;
                }
            }
            // Text rendering mode (ignored for now)
            "Tr" => {}
            "Td" => {
                if op.operands.len() >= 2 {
                    let tx = op.operands[0].as_f64().unwrap_or(0.0);
                    let ty = op.operands[1].as_f64().unwrap_or(0.0);
                    // Td translates the line matrix
                    let new_lm = [
                        state.line_matrix[0], state.line_matrix[1],
                        state.line_matrix[2], state.line_matrix[3],
                        tx * state.line_matrix[0] + ty * state.line_matrix[2] + state.line_matrix[4],
                        tx * state.line_matrix[1] + ty * state.line_matrix[3] + state.line_matrix[5],
                    ];
                    state.line_matrix = new_lm;
                    state.text_matrix = state.line_matrix;
                }
            }
            "TD" => {
                // TD = set leading to -ty, then Td
                if op.operands.len() >= 2 {
                    let tx = op.operands[0].as_f64().unwrap_or(0.0);
                    let ty = op.operands[1].as_f64().unwrap_or(0.0);
                    state.text_leading = -ty;
                    let new_lm = [
                        state.line_matrix[0], state.line_matrix[1],
                        state.line_matrix[2], state.line_matrix[3],
                        tx * state.line_matrix[0] + ty * state.line_matrix[2] + state.line_matrix[4],
                        tx * state.line_matrix[1] + ty * state.line_matrix[3] + state.line_matrix[5],
                    ];
                    state.line_matrix = new_lm;
                    state.text_matrix = state.line_matrix;
                }
            }
            "Tm" => {
                if op.operands.len() >= 6 {
                    let m = extract_matrix(&op.operands);
                    state.text_matrix = m;
                    state.line_matrix = m;
                }
            }
            "T*" => {
                // Move to next line using text_leading
                let leading = if state.text_leading != 0.0 {
                    state.text_leading
                } else {
                    state.font_size
                };
                let tx = 0.0;
                let ty = -leading;
                let new_lm = [
                    state.line_matrix[0], state.line_matrix[1],
                    state.line_matrix[2], state.line_matrix[3],
                    tx * state.line_matrix[0] + ty * state.line_matrix[2] + state.line_matrix[4],
                    tx * state.line_matrix[1] + ty * state.line_matrix[3] + state.line_matrix[5],
                ];
                state.line_matrix = new_lm;
                state.text_matrix = state.line_matrix;
            }
            "Tj" => {
                if let Some(text_bytes) = op.operands.first().and_then(|o| o.as_bytes()) {
                    let text = decode_text(text_bytes, &state.font_name, resources);
                    emit_text(&state, text, &mut elements);
                }
            }
            "TJ" => {
                // Array of strings and positioning adjustments.
                if let Some(arr) = op.operands.first().and_then(|o| o.as_array()) {
                    let mut combined = String::new();
                    for item in arr {
                        match item {
                            PdfObject::String(b) | PdfObject::HexString(b) => {
                                combined.push_str(&decode_text(b, &state.font_name, resources));
                            }
                            _ => {} // Numeric kerning adjustments
                        }
                    }
                    emit_text(&state, combined, &mut elements);
                }
            }
            "'" => {
                // Move to next line and show text (equivalent to T* then Tj).
                let leading = if state.text_leading != 0.0 {
                    state.text_leading
                } else {
                    state.font_size
                };
                let ty = -leading;
                let new_lm = [
                    state.line_matrix[0], state.line_matrix[1],
                    state.line_matrix[2], state.line_matrix[3],
                    ty * state.line_matrix[2] + state.line_matrix[4],
                    ty * state.line_matrix[3] + state.line_matrix[5],
                ];
                state.line_matrix = new_lm;
                state.text_matrix = state.line_matrix;
                if let Some(text_bytes) = op.operands.first().and_then(|o| o.as_bytes()) {
                    let text = decode_text(text_bytes, &state.font_name, resources);
                    emit_text(&state, text, &mut elements);
                }
            }
            "\"" => {
                // Set word spacing, char spacing, move to next line, show text.
                if op.operands.len() >= 3 {
                    state.word_spacing = op.operands[0].as_f64().unwrap_or(0.0);
                    state.char_spacing = op.operands[1].as_f64().unwrap_or(0.0);
                    let leading = if state.text_leading != 0.0 {
                        state.text_leading
                    } else {
                        state.font_size
                    };
                    let ty = -leading;
                    let new_lm = [
                        state.line_matrix[0], state.line_matrix[1],
                        state.line_matrix[2], state.line_matrix[3],
                        ty * state.line_matrix[2] + state.line_matrix[4],
                        ty * state.line_matrix[3] + state.line_matrix[5],
                    ];
                    state.line_matrix = new_lm;
                    state.text_matrix = state.line_matrix;
                    if let Some(text_bytes) = op.operands.get(2).and_then(|o| o.as_bytes()) {
                        let text = decode_text(text_bytes, &state.font_name, resources);
                        emit_text(&state, text, &mut elements);
                    }
                }
            }

            // --- Color operators ---
            "g" => {
                if let Some(gray) = op.operands.first().and_then(|o| o.as_f64()) {
                    state.fill_color = Color::Gray(gray);
                }
            }
            "G" => {
                if let Some(gray) = op.operands.first().and_then(|o| o.as_f64()) {
                    state.stroke_color = Color::Gray(gray);
                }
            }
            "rg" => {
                if op.operands.len() >= 3 {
                    let r = op.operands[0].as_f64().unwrap_or(0.0);
                    let g = op.operands[1].as_f64().unwrap_or(0.0);
                    let b = op.operands[2].as_f64().unwrap_or(0.0);
                    state.fill_color = Color::Rgb(r, g, b);
                }
            }
            "RG" => {
                if op.operands.len() >= 3 {
                    let r = op.operands[0].as_f64().unwrap_or(0.0);
                    let g = op.operands[1].as_f64().unwrap_or(0.0);
                    let b = op.operands[2].as_f64().unwrap_or(0.0);
                    state.stroke_color = Color::Rgb(r, g, b);
                }
            }
            "k" => {
                if op.operands.len() >= 4 {
                    let c = op.operands[0].as_f64().unwrap_or(0.0);
                    let m = op.operands[1].as_f64().unwrap_or(0.0);
                    let y = op.operands[2].as_f64().unwrap_or(0.0);
                    let k = op.operands[3].as_f64().unwrap_or(0.0);
                    state.fill_color = Color::Cmyk(c, m, y, k);
                }
            }
            "K" => {
                if op.operands.len() >= 4 {
                    let c = op.operands[0].as_f64().unwrap_or(0.0);
                    let m = op.operands[1].as_f64().unwrap_or(0.0);
                    let y = op.operands[2].as_f64().unwrap_or(0.0);
                    let k = op.operands[3].as_f64().unwrap_or(0.0);
                    state.stroke_color = Color::Cmyk(c, m, y, k);
                }
            }
            // sc/SC/scn/SCN: set color in current color space
            // Treat as RGB if 3 args, Gray if 1 arg, CMYK if 4 args
            "sc" | "scn" => {
                match op.operands.len() {
                    1 => {
                        let v = op.operands[0].as_f64().unwrap_or(0.0);
                        state.fill_color = Color::Gray(v);
                    }
                    3 => {
                        let r = op.operands[0].as_f64().unwrap_or(0.0);
                        let g = op.operands[1].as_f64().unwrap_or(0.0);
                        let b = op.operands[2].as_f64().unwrap_or(0.0);
                        state.fill_color = Color::Rgb(r, g, b);
                    }
                    4 => {
                        let c = op.operands[0].as_f64().unwrap_or(0.0);
                        let m = op.operands[1].as_f64().unwrap_or(0.0);
                        let y = op.operands[2].as_f64().unwrap_or(0.0);
                        let k = op.operands[3].as_f64().unwrap_or(0.0);
                        state.fill_color = Color::Cmyk(c, m, y, k);
                    }
                    _ => {}
                }
            }
            "SC" | "SCN" => {
                match op.operands.len() {
                    1 => {
                        let v = op.operands[0].as_f64().unwrap_or(0.0);
                        state.stroke_color = Color::Gray(v);
                    }
                    3 => {
                        let r = op.operands[0].as_f64().unwrap_or(0.0);
                        let g = op.operands[1].as_f64().unwrap_or(0.0);
                        let b = op.operands[2].as_f64().unwrap_or(0.0);
                        state.stroke_color = Color::Rgb(r, g, b);
                    }
                    4 => {
                        let c = op.operands[0].as_f64().unwrap_or(0.0);
                        let m = op.operands[1].as_f64().unwrap_or(0.0);
                        let y = op.operands[2].as_f64().unwrap_or(0.0);
                        let k = op.operands[3].as_f64().unwrap_or(0.0);
                        state.stroke_color = Color::Cmyk(c, m, y, k);
                    }
                    _ => {}
                }
            }
            // Color space operators (track but don't fully implement)
            "cs" | "CS" => {}

            // --- Path construction ---
            "m" => {
                if op.operands.len() >= 2 {
                    let x = op.operands[0].as_f64().unwrap_or(0.0);
                    let y = op.operands[1].as_f64().unwrap_or(0.0);
                    current_path.push(PathOp::MoveTo(x, y));
                }
            }
            "l" => {
                if op.operands.len() >= 2 {
                    let x = op.operands[0].as_f64().unwrap_or(0.0);
                    let y = op.operands[1].as_f64().unwrap_or(0.0);
                    current_path.push(PathOp::LineTo(x, y));
                }
            }
            "c" => {
                if op.operands.len() >= 6 {
                    let x1 = op.operands[0].as_f64().unwrap_or(0.0);
                    let y1 = op.operands[1].as_f64().unwrap_or(0.0);
                    let x2 = op.operands[2].as_f64().unwrap_or(0.0);
                    let y2 = op.operands[3].as_f64().unwrap_or(0.0);
                    let x3 = op.operands[4].as_f64().unwrap_or(0.0);
                    let y3 = op.operands[5].as_f64().unwrap_or(0.0);
                    current_path.push(PathOp::CurveTo(x1, y1, x2, y2, x3, y3));
                }
            }
            // v: curve with first control point = current point
            "v" => {
                if op.operands.len() >= 4 {
                    let x2 = op.operands[0].as_f64().unwrap_or(0.0);
                    let y2 = op.operands[1].as_f64().unwrap_or(0.0);
                    let x3 = op.operands[2].as_f64().unwrap_or(0.0);
                    let y3 = op.operands[3].as_f64().unwrap_or(0.0);
                    // Get current point from last path op
                    let (cx, cy) = last_path_point(&current_path);
                    current_path.push(PathOp::CurveTo(cx, cy, x2, y2, x3, y3));
                }
            }
            // y: curve with last control point = endpoint
            "y" => {
                if op.operands.len() >= 4 {
                    let x1 = op.operands[0].as_f64().unwrap_or(0.0);
                    let y1 = op.operands[1].as_f64().unwrap_or(0.0);
                    let x3 = op.operands[2].as_f64().unwrap_or(0.0);
                    let y3 = op.operands[3].as_f64().unwrap_or(0.0);
                    current_path.push(PathOp::CurveTo(x1, y1, x3, y3, x3, y3));
                }
            }
            "h" => current_path.push(PathOp::ClosePath),
            "re" => {
                // Rectangle shorthand: x y w h re
                if op.operands.len() >= 4 {
                    let x = op.operands[0].as_f64().unwrap_or(0.0);
                    let y = op.operands[1].as_f64().unwrap_or(0.0);
                    let w = op.operands[2].as_f64().unwrap_or(0.0);
                    let h = op.operands[3].as_f64().unwrap_or(0.0);
                    current_path.push(PathOp::MoveTo(x, y));
                    current_path.push(PathOp::LineTo(x + w, y));
                    current_path.push(PathOp::LineTo(x + w, y + h));
                    current_path.push(PathOp::LineTo(x, y + h));
                    current_path.push(PathOp::ClosePath);
                }
            }

            // --- Path painting ---
            "S" => {
                // Stroke
                if !current_path.is_empty() {
                    emit_clip_if_pending(&mut pending_clip, &current_path, &state.ctm, &mut elements);
                    let transformed = transform_path_ops(&current_path, &state.ctm);
                    elements.push(ContentElement::Path(PathData {
                        operations: transformed,
                        stroke: Some(make_stroke(&state)),
                        fill: None,
                    }));
                    current_path.clear();
                }
            }
            "s" => {
                // Close and stroke
                current_path.push(PathOp::ClosePath);
                if !current_path.is_empty() {
                    emit_clip_if_pending(&mut pending_clip, &current_path, &state.ctm, &mut elements);
                    let transformed = transform_path_ops(&current_path, &state.ctm);
                    elements.push(ContentElement::Path(PathData {
                        operations: transformed,
                        stroke: Some(make_stroke(&state)),
                        fill: None,
                    }));
                    current_path.clear();
                }
            }
            "f" | "F" => {
                // Fill (non-zero winding)
                if !current_path.is_empty() {
                    emit_clip_if_pending(&mut pending_clip, &current_path, &state.ctm, &mut elements);
                    let transformed = transform_path_ops(&current_path, &state.ctm);
                    elements.push(ContentElement::Path(PathData {
                        operations: transformed,
                        stroke: None,
                        fill: Some(state.fill_color),
                    }));
                    current_path.clear();
                }
            }
            "f*" => {
                // Fill (even-odd rule)
                if !current_path.is_empty() {
                    emit_clip_if_pending(&mut pending_clip, &current_path, &state.ctm, &mut elements);
                    let transformed = transform_path_ops(&current_path, &state.ctm);
                    elements.push(ContentElement::Path(PathData {
                        operations: transformed,
                        stroke: None,
                        fill: Some(state.fill_color),
                    }));
                    current_path.clear();
                }
            }
            "B" => {
                // Fill and stroke (non-zero)
                if !current_path.is_empty() {
                    emit_clip_if_pending(&mut pending_clip, &current_path, &state.ctm, &mut elements);
                    let transformed = transform_path_ops(&current_path, &state.ctm);
                    elements.push(ContentElement::Path(PathData {
                        operations: transformed,
                        stroke: Some(make_stroke(&state)),
                        fill: Some(state.fill_color),
                    }));
                    current_path.clear();
                }
            }
            "B*" => {
                // Fill (even-odd) and stroke
                if !current_path.is_empty() {
                    emit_clip_if_pending(&mut pending_clip, &current_path, &state.ctm, &mut elements);
                    let transformed = transform_path_ops(&current_path, &state.ctm);
                    elements.push(ContentElement::Path(PathData {
                        operations: transformed,
                        stroke: Some(make_stroke(&state)),
                        fill: Some(state.fill_color),
                    }));
                    current_path.clear();
                }
            }
            "b" => {
                // Close, fill and stroke (non-zero)
                current_path.push(PathOp::ClosePath);
                if !current_path.is_empty() {
                    emit_clip_if_pending(&mut pending_clip, &current_path, &state.ctm, &mut elements);
                    let transformed = transform_path_ops(&current_path, &state.ctm);
                    elements.push(ContentElement::Path(PathData {
                        operations: transformed,
                        stroke: Some(make_stroke(&state)),
                        fill: Some(state.fill_color),
                    }));
                    current_path.clear();
                }
            }
            "b*" => {
                // Close, fill (even-odd) and stroke
                current_path.push(PathOp::ClosePath);
                if !current_path.is_empty() {
                    emit_clip_if_pending(&mut pending_clip, &current_path, &state.ctm, &mut elements);
                    let transformed = transform_path_ops(&current_path, &state.ctm);
                    elements.push(ContentElement::Path(PathData {
                        operations: transformed,
                        stroke: Some(make_stroke(&state)),
                        fill: Some(state.fill_color),
                    }));
                    current_path.clear();
                }
            }
            "n" => {
                // End path without painting.
                // If a clip was pending, emit it now.
                if let Some(even_odd) = pending_clip.take() {
                    if !current_path.is_empty() {
                        let transformed = transform_path_ops(&current_path, &state.ctm);
                        elements.push(ContentElement::ClipPath(ClipPathData {
                            operations: transformed,
                            even_odd,
                        }));
                    }
                }
                current_path.clear();
            }
            // Clipping path operators — flag pending, emitted on next paint/n op
            "W" => {
                pending_clip = Some(false);
            }
            "W*" => {
                pending_clip = Some(true);
            }

            // --- XObject (Form XObject / Image) ---
            "Do" => {
                if let Some(name) = op.operands.first().and_then(|o| o.as_name()) {
                    if let Some(xobj) = resources.xobject_streams.get(name) {
                        // Save current CTM, apply XObject's matrix, interpret, restore
                        let saved_ctm = state.ctm;
                        state.ctm = multiply_matrix(&saved_ctm, &xobj.matrix);
                        // Use XObject's own resources if available, otherwise inherit page resources
                        let sub_resources = match &xobj.resources {
                            Some(r) => r.as_ref(),
                            None => resources,
                        };
                        if let Ok(sub_elements) = interpret_content_stream_impl(
                            &xobj.stream, sub_resources, depth + 1,
                        ) {
                            elements.extend(sub_elements);
                        }
                        state.ctm = saved_ctm;
                    }
                }
            }
            // Inline image / marked content (skip)
            "BI" | "ID" | "EI" | "BMC" | "BDC" | "EMC" | "MP" | "DP" | "sh" => {}

            _ => {
                // Unknown operator — skip.
            }
        }
    }

    Ok(elements)
}

/// Maximum number of operators in a single content stream.
const MAX_OPERATORS: usize = 1_000_000;

/// Tokenize a content stream into a sequence of operators with their operands.
fn tokenize_content_stream(data: &[u8]) -> Result<Vec<Operator>, PdfError> {
    let mut operators = Vec::new();
    let mut operands: Vec<PdfObject> = Vec::new();
    let mut pos = 0;

    while pos < data.len() {
        pos = skip_content_ws(data, pos);
        if pos >= data.len() {
            break;
        }

        match data[pos] {
            // Number
            b'+' | b'-' | b'.' | b'0'..=b'9' => {
                let (obj, new_pos) = parse_content_number(data, pos)?;
                operands.push(obj);
                pos = new_pos;
            }
            // Name
            b'/' => {
                let (name, new_pos) = parse_content_name(data, pos);
                operands.push(PdfObject::Name(name));
                pos = new_pos;
            }
            // String
            b'(' => {
                let (s, new_pos) = parse_content_string(data, pos)?;
                operands.push(PdfObject::String(s));
                pos = new_pos;
            }
            // Hex string
            b'<' => {
                let (s, new_pos) = parse_content_hex_string(data, pos)?;
                operands.push(PdfObject::HexString(s));
                pos = new_pos;
            }
            // Array
            b'[' => {
                let (arr, new_pos) = parse_content_array(data, pos)?;
                operands.push(arr);
                pos = new_pos;
            }
            // Operator keyword (alphabetic)
            b'a'..=b'z' | b'A'..=b'Z' | b'\'' | b'"' | b'*' => {
                let start = pos;
                while pos < data.len()
                    && (data[pos].is_ascii_alphabetic()
                        || data[pos] == b'*'
                        || data[pos] == b'\''
                        || data[pos] == b'"')
                {
                    pos += 1;
                }
                let name = String::from_utf8_lossy(&data[start..pos]).into_owned();

                // "true", "false", "null" are operands not operators.
                match name.as_str() {
                    "true" => operands.push(PdfObject::Boolean(true)),
                    "false" => operands.push(PdfObject::Boolean(false)),
                    "null" => operands.push(PdfObject::Null),
                    "BI" => {
                        // Inline image: BI <dict pairs> ID <binary data> EI
                        // Skip everything until we find "EI" preceded by whitespace.
                        // First skip to "ID"
                        while pos < data.len() {
                            pos = skip_content_ws(data, pos);
                            if pos + 2 <= data.len()
                                && data[pos] == b'I'
                                && data[pos + 1] == b'D'
                                && (pos + 2 >= data.len() || !data[pos + 2].is_ascii_alphabetic())
                            {
                                pos += 2;
                                // Skip single whitespace/newline after ID
                                if pos < data.len() && (data[pos] == b' ' || data[pos] == b'\n' || data[pos] == b'\r') {
                                    pos += 1;
                                }
                                break;
                            }
                            // Skip key-value pair tokens
                            if pos < data.len() {
                                if data[pos] == b'/' {
                                    // Name
                                    pos += 1;
                                    while pos < data.len() && !data[pos].is_ascii_whitespace()
                                        && !matches!(data[pos], b'/' | b'<' | b'>' | b'[' | b']')
                                    {
                                        pos += 1;
                                    }
                                } else if data[pos].is_ascii_digit() || data[pos] == b'-' || data[pos] == b'+' || data[pos] == b'.' {
                                    while pos < data.len() && (data[pos].is_ascii_digit() || data[pos] == b'.') {
                                        pos += 1;
                                    }
                                } else if data[pos].is_ascii_alphabetic() {
                                    while pos < data.len() && data[pos].is_ascii_alphabetic() {
                                        pos += 1;
                                    }
                                } else {
                                    pos += 1;
                                }
                            }
                        }
                        // Now skip binary data until EI
                        // EI must be preceded by whitespace and followed by whitespace/EOF
                        while pos < data.len() {
                            if pos + 2 <= data.len()
                                && data[pos] == b'E'
                                && data[pos + 1] == b'I'
                                && (pos == 0 || data[pos - 1].is_ascii_whitespace())
                                && (pos + 2 >= data.len() || data[pos + 2].is_ascii_whitespace())
                            {
                                pos += 2; // skip "EI"
                                break;
                            }
                            pos += 1;
                        }
                        // Push BI as operator (will be ignored in interpreter)
                        operators.push(Operator {
                            name,
                            operands: std::mem::take(&mut operands),
                        });
                    }
                    _ => {
                        operators.push(Operator {
                            name,
                            operands: std::mem::take(&mut operands),
                        });
                        if operators.len() > MAX_OPERATORS {
                            return Err(PdfError::Parse(format!(
                                "content stream exceeds {MAX_OPERATORS} operator limit"
                            )));
                        }
                    }
                }
            }
            _ => {
                pos += 1; // skip unknown byte
            }
        }
    }

    Ok(operators)
}

fn parse_content_number(data: &[u8], pos: usize) -> Result<(PdfObject, usize), PdfError> {
    let mut end = pos;
    if end < data.len() && (data[end] == b'+' || data[end] == b'-') {
        end += 1;
    }
    let mut has_dot = false;
    while end < data.len() && (data[end].is_ascii_digit() || data[end] == b'.') {
        if data[end] == b'.' {
            has_dot = true;
        }
        end += 1;
    }
    let s = std::str::from_utf8(&data[pos..end])
        .map_err(|_| PdfError::Parse("invalid number bytes".into()))?;
    if has_dot {
        let val: f64 = s
            .parse()
            .map_err(|_| PdfError::Parse(format!("invalid real: {s}")))?;
        Ok((PdfObject::Real(val), end))
    } else {
        let val: i64 = s
            .parse()
            .map_err(|_| PdfError::Parse(format!("invalid integer: {s}")))?;
        Ok((PdfObject::Integer(val), end))
    }
}

fn parse_content_name(data: &[u8], mut pos: usize) -> (String, usize) {
    pos += 1; // skip '/'
    let start = pos;
    while pos < data.len()
        && !data[pos].is_ascii_whitespace()
        && !matches!(
            data[pos],
            b'/' | b'<' | b'>' | b'[' | b']' | b'(' | b')' | b'{' | b'}'
        )
    {
        pos += 1;
    }
    let name = String::from_utf8_lossy(&data[start..pos]).into_owned();
    (name, pos)
}

fn parse_content_string(data: &[u8], mut pos: usize) -> Result<(Vec<u8>, usize), PdfError> {
    pos += 1; // skip '('
    let mut result = Vec::new();
    let mut depth = 1u32;

    while pos < data.len() && depth > 0 {
        match data[pos] {
            b'(' => {
                depth += 1;
                result.push(b'(');
            }
            b')' => {
                depth -= 1;
                if depth > 0 {
                    result.push(b')');
                }
            }
            b'\\' => {
                pos += 1;
                if pos < data.len() {
                    match data[pos] {
                        b'n' => result.push(b'\n'),
                        b'r' => result.push(b'\r'),
                        b't' => result.push(b'\t'),
                        b'(' => result.push(b'('),
                        b')' => result.push(b')'),
                        b'\\' => result.push(b'\\'),
                        b'0'..=b'7' => {
                            // Octal escape
                            let mut val = (data[pos] - b'0') as u8;
                            for _ in 0..2 {
                                if pos + 1 < data.len() && data[pos + 1] >= b'0' && data[pos + 1] <= b'7'
                                {
                                    pos += 1;
                                    val = val * 8 + (data[pos] - b'0');
                                } else {
                                    break;
                                }
                            }
                            result.push(val);
                        }
                        other => result.push(other),
                    }
                }
            }
            b => result.push(b),
        }
        pos += 1;
    }
    Ok((result, pos))
}

fn parse_content_hex_string(data: &[u8], mut pos: usize) -> Result<(Vec<u8>, usize), PdfError> {
    pos += 1; // skip '<'
    let mut hex = Vec::new();
    while pos < data.len() && data[pos] != b'>' {
        if !data[pos].is_ascii_whitespace() {
            hex.push(data[pos]);
        }
        pos += 1;
    }
    if pos < data.len() {
        pos += 1; // skip '>'
    }
    if hex.len() % 2 != 0 {
        hex.push(b'0');
    }
    let bytes: Vec<u8> = hex
        .chunks(2)
        .filter_map(|pair| {
            let s = std::str::from_utf8(pair).ok()?;
            u8::from_str_radix(s, 16).ok()
        })
        .collect();
    Ok((bytes, pos))
}

fn parse_content_array(data: &[u8], mut pos: usize) -> Result<(PdfObject, usize), PdfError> {
    pos += 1; // skip '['
    pos = skip_content_ws(data, pos);
    let mut items = Vec::new();

    while pos < data.len() && data[pos] != b']' {
        match data[pos] {
            b'+' | b'-' | b'.' | b'0'..=b'9' => {
                let (obj, new_pos) = parse_content_number(data, pos)?;
                items.push(obj);
                pos = new_pos;
            }
            b'(' => {
                let (s, new_pos) = parse_content_string(data, pos)?;
                items.push(PdfObject::String(s));
                pos = new_pos;
            }
            b'<' => {
                let (s, new_pos) = parse_content_hex_string(data, pos)?;
                items.push(PdfObject::HexString(s));
                pos = new_pos;
            }
            b'/' => {
                let (name, new_pos) = parse_content_name(data, pos);
                items.push(PdfObject::Name(name));
                pos = new_pos;
            }
            _ => pos += 1,
        }
        pos = skip_content_ws(data, pos);
    }
    if pos < data.len() {
        pos += 1; // skip ']'
    }
    Ok((PdfObject::Array(items), pos))
}

fn skip_content_ws(data: &[u8], mut pos: usize) -> usize {
    while pos < data.len() {
        match data[pos] {
            b' ' | b'\t' | b'\r' | b'\n' | 0x0C | 0x00 => pos += 1,
            b'%' => {
                while pos < data.len() && data[pos] != b'\n' && data[pos] != b'\r' {
                    pos += 1;
                }
            }
            _ => break,
        }
    }
    pos
}

/// Decode text bytes using CMap, font encoding, or fallback.
fn decode_text(
    bytes: &[u8],
    font_name: &str,
    resources: &PageResources,
) -> String {
    // 1. Try ToUnicode CMap first (most accurate)
    if let Some(cmap) = resources.font_cmaps.get(font_name) {
        return cmap.decode_bytes(bytes);
    }
    // 2. Try font encoding (WinAnsiEncoding, MacRomanEncoding)
    if let Some(enc) = resources.font_encodings.get(font_name) {
        return decode_with_encoding(bytes, enc);
    }
    // 3. Fallback to PDFDocEncoding
    decode_pdf_text(bytes)
}

/// Decode PDF text bytes to a String (fallback when no CMap is available).
fn decode_pdf_text(bytes: &[u8]) -> String {
    // Check for UTF-16BE BOM.
    if bytes.len() >= 2 && bytes[0] == 0xFE && bytes[1] == 0xFF {
        let chars: Vec<u16> = bytes[2..]
            .chunks(2)
            .filter_map(|chunk| {
                if chunk.len() == 2 {
                    Some(u16::from_be_bytes([chunk[0], chunk[1]]))
                } else {
                    None
                }
            })
            .collect();
        String::from_utf16_lossy(&chars)
    } else {
        // PDFDocEncoding (roughly Latin-1 for the common range).
        bytes.iter().map(|&b| b as char).collect()
    }
}

fn extract_matrix(operands: &[PdfObject]) -> [f64; 6] {
    let mut m = [0.0f64; 6];
    for (i, op) in operands.iter().take(6).enumerate() {
        m[i] = op.as_f64().unwrap_or(0.0);
    }
    m
}

fn multiply_matrix(a: &[f64; 6], b: &[f64; 6]) -> [f64; 6] {
    [
        a[0] * b[0] + a[1] * b[2],
        a[0] * b[1] + a[1] * b[3],
        a[2] * b[0] + a[3] * b[2],
        a[2] * b[1] + a[3] * b[3],
        a[4] * b[0] + a[5] * b[2] + b[4],
        a[4] * b[1] + a[5] * b[3] + b[5],
    ]
}

fn transform_point(ctm: &[f64; 6], tm: &[f64; 6]) -> (f64, f64) {
    // Combine text matrix with CTM.
    let combined = multiply_matrix(tm, ctm);
    (combined[4], combined[5])
}

/// Transform a single point through a CTM.
fn ctm_transform(ctm: &[f64; 6], x: f64, y: f64) -> (f64, f64) {
    (
        x * ctm[0] + y * ctm[2] + ctm[4],
        x * ctm[1] + y * ctm[3] + ctm[5],
    )
}

/// Transform all path operations through a CTM.
fn transform_path_ops(ops: &[PathOp], ctm: &[f64; 6]) -> Vec<PathOp> {
    // If CTM is identity, skip transformation
    if (ctm[0] - 1.0).abs() < 1e-10
        && ctm[1].abs() < 1e-10
        && ctm[2].abs() < 1e-10
        && (ctm[3] - 1.0).abs() < 1e-10
        && ctm[4].abs() < 1e-10
        && ctm[5].abs() < 1e-10
    {
        return ops.to_vec();
    }
    ops.iter()
        .map(|op| match op {
            PathOp::MoveTo(x, y) => {
                let (tx, ty) = ctm_transform(ctm, *x, *y);
                PathOp::MoveTo(tx, ty)
            }
            PathOp::LineTo(x, y) => {
                let (tx, ty) = ctm_transform(ctm, *x, *y);
                PathOp::LineTo(tx, ty)
            }
            PathOp::CurveTo(x1, y1, x2, y2, x3, y3) => {
                let (tx1, ty1) = ctm_transform(ctm, *x1, *y1);
                let (tx2, ty2) = ctm_transform(ctm, *x2, *y2);
                let (tx3, ty3) = ctm_transform(ctm, *x3, *y3);
                PathOp::CurveTo(tx1, ty1, tx2, ty2, tx3, ty3)
            }
            PathOp::ClosePath => PathOp::ClosePath,
        })
        .collect()
}

/// Emit a ClipPath element if a clip is pending.
fn emit_clip_if_pending(
    pending_clip: &mut Option<bool>,
    current_path: &[PathOp],
    ctm: &[f64; 6],
    elements: &mut Vec<ContentElement>,
) {
    if let Some(even_odd) = pending_clip.take() {
        if !current_path.is_empty() {
            let transformed = transform_path_ops(current_path, ctm);
            elements.push(ContentElement::ClipPath(ClipPathData {
                operations: transformed,
                even_odd,
            }));
        }
    }
}

/// Get the current point from the last path operation (for `v` operator).
fn last_path_point(path: &[PathOp]) -> (f64, f64) {
    for op in path.iter().rev() {
        match op {
            PathOp::MoveTo(x, y) | PathOp::LineTo(x, y) => return (*x, *y),
            PathOp::CurveTo(_, _, _, _, x, y) => return (*x, *y),
            PathOp::ClosePath => continue,
        }
    }
    (0.0, 0.0)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_tokenize_basic() {
        let data = b"BT /F1 12 Tf 100 700 Td (Hello World) Tj ET";
        let ops = tokenize_content_stream(data).unwrap();

        assert_eq!(ops.len(), 5);
        assert_eq!(ops[0].name, "BT");
        assert_eq!(ops[1].name, "Tf");
        assert_eq!(ops[2].name, "Td");
        assert_eq!(ops[3].name, "Tj");
        assert_eq!(ops[4].name, "ET");
    }

    #[test]
    fn test_interpret_text() {
        let data = b"BT /F1 12 Tf 100 700 Td (Hello) Tj ET";
        let elements = interpret_content_stream(data, 792.0).unwrap();

        assert_eq!(elements.len(), 1);
        match &elements[0] {
            ContentElement::Text(span) => {
                assert_eq!(span.text, "Hello");
                assert_eq!(span.font_size, 12.0);
                assert!((span.x - 100.0).abs() < 0.01);
                assert!((span.y - 700.0).abs() < 0.01); // raw PDF coords
            }
            _ => panic!("expected text element"),
        }
    }

    #[test]
    fn test_interpret_tj_array() {
        let data = b"BT /F1 10 Tf 50 500 Td [(Hello) -100 ( World)] TJ ET";
        let elements = interpret_content_stream(data, 792.0).unwrap();

        assert_eq!(elements.len(), 1);
        match &elements[0] {
            ContentElement::Text(span) => {
                assert_eq!(span.text, "Hello World");
            }
            _ => panic!("expected text element"),
        }
    }

    #[test]
    fn test_interpret_path() {
        let data = b"100 100 200 50 re S";
        let elements = interpret_content_stream(data, 792.0).unwrap();

        assert_eq!(elements.len(), 1);
        match &elements[0] {
            ContentElement::Path(path) => {
                assert_eq!(path.operations.len(), 5); // moveTo + 3 lineTo + close
                assert!(path.stroke.is_some());
                assert!(path.fill.is_none());
            }
            _ => panic!("expected path element"),
        }
    }

    #[test]
    fn test_color_operators() {
        let data = b"BT /F1 12 Tf 1 0 0 rg 100 700 Td (Red) Tj ET";
        let elements = interpret_content_stream(data, 792.0).unwrap();

        match &elements[0] {
            ContentElement::Text(span) => match span.fill_color {
                Color::Rgb(r, g, b) => {
                    assert!((r - 1.0).abs() < 0.01);
                    assert!(g.abs() < 0.01);
                    assert!(b.abs() < 0.01);
                }
                _ => panic!("expected RGB color"),
            },
            _ => panic!("expected text"),
        }
    }

    #[test]
    fn test_decode_utf16() {
        let bytes = [0xFE, 0xFF, 0x00, 0x48, 0x00, 0x69]; // "Hi" in UTF-16BE
        let text = decode_pdf_text(&bytes);
        assert_eq!(text, "Hi");
    }
}
