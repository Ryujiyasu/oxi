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

#[derive(Serialize)]
struct LayoutElementJs {
    x: f32,
    y: f32,
    width: f32,
    height: f32,
    kind: String,
    text: Option<String>,
    font_size: Option<f32>,
    bold: Option<bool>,
    italic: Option<bool>,
    color: Option<String>,
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
                            text, font_size, bold, italic, color, ..
                        } => LayoutElementJs {
                            x: elem.x, y: elem.y, width: elem.width, height: elem.height,
                            kind: "text".into(),
                            text: Some(text),
                            font_size: Some(font_size),
                            bold: Some(bold),
                            italic: Some(italic),
                            color,
                            x1: None, y1: None, x2: None, y2: None,
                        },
                        oxidocs_core::layout::LayoutContent::Image { .. } => LayoutElementJs {
                            x: elem.x, y: elem.y, width: elem.width, height: elem.height,
                            kind: "image".into(),
                            text: None, font_size: None, bold: None, italic: None, color: None,
                            x1: None, y1: None, x2: None, y2: None,
                        },
                        oxidocs_core::layout::LayoutContent::TableBorder { x1, y1, x2, y2 } => LayoutElementJs {
                            x: elem.x, y: elem.y, width: elem.width, height: elem.height,
                            kind: "border".into(),
                            text: None, font_size: None, bold: None, italic: None, color: None,
                            x1: Some(x1), y1: Some(y1), x2: Some(x2), y2: Some(y2),
                        },
                    }
                }).collect(),
            }
        }).collect(),
    };

    serde_wasm_bindgen::to_value(&js_result).map_err(|e| JsError::new(&e.to_string()))
}
