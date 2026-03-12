use wasm_bindgen::prelude::*;
use serde::Serialize;

#[wasm_bindgen(start)]
pub fn init() {
    console_error_panic_hook::set_once();
}

#[wasm_bindgen]
pub fn parse_document(data: &[u8]) -> Result<JsValue, JsError> {
    let doc = oxidocs_core::parse_docx(data)
        .map_err(|e| JsError::new(&e.to_string()))?;
    serde_wasm_bindgen::to_value(&doc).map_err(|e| JsError::new(&e.to_string()))
}

#[wasm_bindgen]
pub fn parse_spreadsheet(data: &[u8]) -> Result<JsValue, JsError> {
    let workbook = oxicells_core::parse_xlsx(data)
        .map_err(|e| JsError::new(&e.to_string()))?;
    serde_wasm_bindgen::to_value(&workbook).map_err(|e| JsError::new(&e.to_string()))
}

#[wasm_bindgen]
pub fn parse_presentation(data: &[u8]) -> Result<JsValue, JsError> {
    let pres = oxislides_core::parse_pptx(data)
        .map_err(|e| JsError::new(&e.to_string()))?;
    serde_wasm_bindgen::to_value(&pres).map_err(|e| JsError::new(&e.to_string()))
}

#[derive(Serialize)]
struct LayoutElementJs {
    x: f32,
    y: f32,
    width: f32,
    height: f32,
    kind: String,
    // Text fields
    text: Option<String>,
    font_size: Option<f32>,
    bold: Option<bool>,
    italic: Option<bool>,
    underline: Option<bool>,
    strikethrough: Option<bool>,
    color: Option<String>,
    highlight: Option<String>,
    // Image fields
    image_data: Option<String>,  // base64-encoded
    content_type: Option<String>,
    // Border fields
    x1: Option<f32>,
    y1: Option<f32>,
    x2: Option<f32>,
    y2: Option<f32>,
}

#[derive(Serialize)]
struct LayoutPageJs {
    width: f32,
    height: f32,
    elements: Vec<LayoutElementJs>,
}

#[derive(Serialize)]
struct LayoutResultJs {
    pages: Vec<LayoutPageJs>,
}

fn base64_encode(data: &[u8]) -> String {
    const CHARS: &[u8] = b"ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/";
    let mut result = String::with_capacity((data.len() + 2) / 3 * 4);
    for chunk in data.chunks(3) {
        let b0 = chunk[0] as u32;
        let b1 = if chunk.len() > 1 { chunk[1] as u32 } else { 0 };
        let b2 = if chunk.len() > 2 { chunk[2] as u32 } else { 0 };
        let triple = (b0 << 16) | (b1 << 8) | b2;
        result.push(CHARS[((triple >> 18) & 0x3F) as usize] as char);
        result.push(CHARS[((triple >> 12) & 0x3F) as usize] as char);
        if chunk.len() > 1 {
            result.push(CHARS[((triple >> 6) & 0x3F) as usize] as char);
        } else {
            result.push('=');
        }
        if chunk.len() > 2 {
            result.push(CHARS[(triple & 0x3F) as usize] as char);
        } else {
            result.push('=');
        }
    }
    result
}

#[wasm_bindgen]
pub fn layout_document(data: &[u8]) -> Result<JsValue, JsError> {
    let doc = oxidocs_core::parse_docx(data)
        .map_err(|e| JsError::new(&e.to_string()))?;

    let engine = oxidocs_core::layout::LayoutEngine::new();
    let result = engine.layout(&doc);

    let js_result = LayoutResultJs {
        pages: result.pages.into_iter().map(|page| {
            LayoutPageJs {
                width: page.width,
                height: page.height,
                elements: page.elements.into_iter().map(|elem| {
                    match elem.content {
                        oxidocs_core::layout::LayoutContent::Text {
                            text, font_size, bold, italic, underline, strikethrough, color, highlight, ..
                        } => LayoutElementJs {
                            x: elem.x, y: elem.y, width: elem.width, height: elem.height,
                            kind: "text".into(),
                            text: Some(text),
                            font_size: Some(font_size),
                            bold: Some(bold),
                            italic: Some(italic),
                            underline: Some(underline),
                            strikethrough: Some(strikethrough),
                            color,
                            highlight,
                            image_data: None, content_type: None,
                            x1: None, y1: None, x2: None, y2: None,
                        },
                        oxidocs_core::layout::LayoutContent::Image { data, content_type } => {
                            let b64 = if !data.is_empty() { Some(base64_encode(&data)) } else { None };
                            LayoutElementJs {
                                x: elem.x, y: elem.y, width: elem.width, height: elem.height,
                                kind: "image".into(),
                                text: None, font_size: None, bold: None, italic: None,
                                underline: None, strikethrough: None, color: None, highlight: None,
                                image_data: b64,
                                content_type,
                                x1: None, y1: None, x2: None, y2: None,
                            }
                        },
                        oxidocs_core::layout::LayoutContent::TableBorder { x1, y1, x2, y2 } => LayoutElementJs {
                            x: elem.x, y: elem.y, width: elem.width, height: elem.height,
                            kind: "border".into(),
                            text: None, font_size: None, bold: None, italic: None,
                            underline: None, strikethrough: None, color: None, highlight: None,
                            image_data: None, content_type: None,
                            x1: Some(x1), y1: Some(y1), x2: Some(x2), y2: Some(y2),
                        },
                    }
                }).collect(),
            }
        }).collect(),
    };

    serde_wasm_bindgen::to_value(&js_result).map_err(|e| JsError::new(&e.to_string()))
}
