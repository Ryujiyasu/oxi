//! PDF content stream parser.
//!
//! A content stream is a sequence of operators with their operands that
//! describe text, graphics, and images on a page.

use std::collections::HashMap;

use crate::error::PdfError;
use crate::ir::*;
use super::object::PdfObject;
use super::cmap::CMap;

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
    // Color state
    fill_color: Color,
    stroke_color: Color,
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
            fill_color: Color::Gray(0.0), // black
            stroke_color: Color::Gray(0.0),
            ctm: [1.0, 0.0, 0.0, 1.0, 0.0, 0.0],
        }
    }
}

/// Interpret a content stream and extract content elements.
pub fn interpret_content_stream(
    stream_data: &[u8],
    page_height: f64,
) -> Result<Vec<ContentElement>, PdfError> {
    interpret_content_stream_with_cmaps(stream_data, page_height, &HashMap::new())
}

/// Interpret a content stream with font-specific CMap decoders.
/// The `font_cmaps` map font names (e.g. "/F1") to their ToUnicode CMaps.
pub fn interpret_content_stream_with_cmaps(
    stream_data: &[u8],
    page_height: f64,
    font_cmaps: &HashMap<String, CMap>,
) -> Result<Vec<ContentElement>, PdfError> {
    let operators = tokenize_content_stream(stream_data)?;
    let mut elements = Vec::new();
    let mut state = GraphicsState::default();
    let mut state_stack: Vec<GraphicsState> = Vec::new();
    let mut current_path: Vec<PathOp> = Vec::new();

    for op in &operators {
        match op.name.as_str() {
            // --- Graphics state ---
            "q" => state_stack.push(state.clone()),
            "Q" => {
                if let Some(saved) = state_stack.pop() {
                    state = saved;
                }
            }
            "cm" => {
                if op.operands.len() >= 6 {
                    let m = extract_matrix(&op.operands);
                    state.ctm = multiply_matrix(&state.ctm, &m);
                }
            }

            // --- Text operators ---
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
            "Td" => {
                if op.operands.len() >= 2 {
                    let tx = op.operands[0].as_f64().unwrap_or(0.0);
                    let ty = op.operands[1].as_f64().unwrap_or(0.0);
                    state.line_matrix[4] += tx;
                    state.line_matrix[5] += ty;
                    state.text_matrix = state.line_matrix;
                }
            }
            "TD" => {
                // Same as Td but also sets leading.
                if op.operands.len() >= 2 {
                    let tx = op.operands[0].as_f64().unwrap_or(0.0);
                    let ty = op.operands[1].as_f64().unwrap_or(0.0);
                    state.line_matrix[4] += tx;
                    state.line_matrix[5] += ty;
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
                // Move to next line (uses text leading, default ~font_size).
                state.line_matrix[5] -= state.font_size;
                state.text_matrix = state.line_matrix;
            }
            "Tj" => {
                if let Some(text_bytes) = op.operands.first().and_then(|o| o.as_bytes()) {
                    let text = decode_with_cmap(text_bytes, &state.font_name, font_cmaps);
                    if !text.trim().is_empty() {
                        let (x, y) = transform_point(&state.ctm, &state.text_matrix);
                        elements.push(ContentElement::Text(TextSpan {
                            x,
                            y: page_height - y, // PDF y=0 is bottom
                            text,
                            font_name: state.font_name.clone(),
                            font_size: state.font_size,
                            fill_color: state.fill_color,
                        }));
                    }
                }
            }
            "TJ" => {
                // Array of strings and positioning adjustments.
                if let Some(arr) = op.operands.first().and_then(|o| o.as_array()) {
                    let mut combined = String::new();
                    for item in arr {
                        match item {
                            PdfObject::String(b) | PdfObject::HexString(b) => {
                                combined.push_str(&decode_with_cmap(b, &state.font_name, font_cmaps));
                            }
                            _ => {} // Numeric kerning adjustments, skip for now.
                        }
                    }
                    if !combined.trim().is_empty() {
                        let (x, y) = transform_point(&state.ctm, &state.text_matrix);
                        elements.push(ContentElement::Text(TextSpan {
                            x,
                            y: page_height - y,
                            text: combined,
                            font_name: state.font_name.clone(),
                            font_size: state.font_size,
                            fill_color: state.fill_color,
                        }));
                    }
                }
            }
            "'" => {
                // Move to next line and show text.
                state.line_matrix[5] -= state.font_size;
                state.text_matrix = state.line_matrix;
                if let Some(text_bytes) = op.operands.first().and_then(|o| o.as_bytes()) {
                    let text = decode_with_cmap(text_bytes, &state.font_name, font_cmaps);
                    if !text.trim().is_empty() {
                        let (x, y) = transform_point(&state.ctm, &state.text_matrix);
                        elements.push(ContentElement::Text(TextSpan {
                            x,
                            y: page_height - y,
                            text,
                            font_name: state.font_name.clone(),
                            font_size: state.font_size,
                            fill_color: state.fill_color,
                        }));
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

            // --- Path construction ---
            "m" => {
                if op.operands.len() >= 2 {
                    let x = op.operands[0].as_f64().unwrap_or(0.0);
                    let y = op.operands[1].as_f64().unwrap_or(0.0);
                    current_path.push(PathOp::MoveTo(x, page_height - y));
                }
            }
            "l" => {
                if op.operands.len() >= 2 {
                    let x = op.operands[0].as_f64().unwrap_or(0.0);
                    let y = op.operands[1].as_f64().unwrap_or(0.0);
                    current_path.push(PathOp::LineTo(x, page_height - y));
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
                    current_path.push(PathOp::CurveTo(
                        x1,
                        page_height - y1,
                        x2,
                        page_height - y2,
                        x3,
                        page_height - y3,
                    ));
                }
            }
            "h" => current_path.push(PathOp::ClosePath),
            "re" => {
                // Rectangle shorthand: x y w h
                if op.operands.len() >= 4 {
                    let x = op.operands[0].as_f64().unwrap_or(0.0);
                    let y = op.operands[1].as_f64().unwrap_or(0.0);
                    let w = op.operands[2].as_f64().unwrap_or(0.0);
                    let h = op.operands[3].as_f64().unwrap_or(0.0);
                    let fy = page_height - y;
                    current_path.push(PathOp::MoveTo(x, fy));
                    current_path.push(PathOp::LineTo(x + w, fy));
                    current_path.push(PathOp::LineTo(x + w, fy - h));
                    current_path.push(PathOp::LineTo(x, fy - h));
                    current_path.push(PathOp::ClosePath);
                }
            }

            // --- Path painting ---
            "S" => {
                // Stroke
                if !current_path.is_empty() {
                    elements.push(ContentElement::Path(PathData {
                        operations: std::mem::take(&mut current_path),
                        stroke: Some(StrokeStyle {
                            color: state.stroke_color,
                            width: 1.0,
                            line_cap: LineCap::Butt,
                            line_join: LineJoin::Miter,
                        }),
                        fill: None,
                    }));
                }
            }
            "f" | "F" => {
                // Fill
                if !current_path.is_empty() {
                    elements.push(ContentElement::Path(PathData {
                        operations: std::mem::take(&mut current_path),
                        stroke: None,
                        fill: Some(state.fill_color),
                    }));
                }
            }
            "B" => {
                // Fill and stroke
                if !current_path.is_empty() {
                    elements.push(ContentElement::Path(PathData {
                        operations: std::mem::take(&mut current_path),
                        stroke: Some(StrokeStyle {
                            color: state.stroke_color,
                            width: 1.0,
                            line_cap: LineCap::Butt,
                            line_join: LineJoin::Miter,
                        }),
                        fill: Some(state.fill_color),
                    }));
                }
            }
            "n" => {
                // End path without painting (clipping path).
                current_path.clear();
            }

            _ => {
                // Unknown operator — skip for now.
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

/// Decode text bytes using CMap if available for the font, otherwise fallback.
fn decode_with_cmap(
    bytes: &[u8],
    font_name: &str,
    font_cmaps: &HashMap<String, CMap>,
) -> String {
    if let Some(cmap) = font_cmaps.get(font_name) {
        cmap.decode_bytes(bytes)
    } else {
        decode_pdf_text(bytes)
    }
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
                assert!((span.y - 92.0).abs() < 0.01); // 792 - 700
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
