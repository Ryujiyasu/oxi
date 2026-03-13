use wasm_bindgen::prelude::*;
use serde::{Deserialize, Serialize};

#[wasm_bindgen(start)]
pub fn init() {
    console_error_panic_hook::set_once();
}

/// Create a blank .docx file and return it as bytes.
/// Can be used to create a new document from scratch.
#[wasm_bindgen]
pub fn create_blank_docx() -> Vec<u8> {
    oxidocs_core::create_blank_docx()
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

/// A single text edit operation from JavaScript.
#[derive(Deserialize)]
struct JsTextEdit {
    paragraph_index: usize,
    run_index: usize,
    new_text: String,
}

/// Edit a .docx file and return the modified bytes.
///
/// `data`: original .docx bytes
/// `edits`: JS array of `{paragraph_index, run_index, new_text}` objects
///
/// Returns the modified .docx as `Uint8Array`.
#[wasm_bindgen]
pub fn edit_docx(data: &[u8], edits: JsValue) -> Result<Vec<u8>, JsError> {
    let js_edits: Vec<JsTextEdit> = serde_wasm_bindgen::from_value(edits)
        .map_err(|e| JsError::new(&e.to_string()))?;

    let mut editor = oxidocs_core::DocxEditor::new(data)
        .map_err(|e| JsError::new(&e.to_string()))?;

    let edits: Vec<oxidocs_core::editor::TextEdit> = js_edits
        .into_iter()
        .map(|e| oxidocs_core::editor::TextEdit {
            paragraph_index: e.paragraph_index,
            run_index: e.run_index,
            new_text: e.new_text,
        })
        .collect();

    editor.apply_edits(&edits);

    editor
        .save()
        .map_err(|e| JsError::new(&e.to_string()))
}

/// A single cell edit operation from JavaScript.
#[derive(Deserialize)]
struct JsCellEdit {
    sheet_index: usize,
    row: u32,
    col: u32,
    new_value: String,
}

/// Edit a .xlsx file and return the modified bytes.
#[wasm_bindgen]
pub fn edit_xlsx(data: &[u8], edits: JsValue) -> Result<Vec<u8>, JsError> {
    let js_edits: Vec<JsCellEdit> = serde_wasm_bindgen::from_value(edits)
        .map_err(|e| JsError::new(&e.to_string()))?;

    let mut editor = oxicells_core::XlsxEditor::new(data)
        .map_err(|e| JsError::new(&e.to_string()))?;

    let edits: Vec<oxicells_core::editor::CellEdit> = js_edits
        .into_iter()
        .map(|e| oxicells_core::editor::CellEdit {
            sheet_index: e.sheet_index,
            row: e.row,
            col: e.col,
            new_value: e.new_value,
        })
        .collect();

    editor.apply_edits(&edits);

    editor
        .save()
        .map_err(|e| JsError::new(&e.to_string()))
}

/// A single slide text edit operation from JavaScript.
#[derive(Deserialize)]
struct JsSlideTextEdit {
    slide_index: usize,
    shape_index: usize,
    paragraph_index: usize,
    run_index: usize,
    new_text: String,
}

/// Edit a .pptx file and return the modified bytes.
#[wasm_bindgen]
pub fn edit_pptx(data: &[u8], edits: JsValue) -> Result<Vec<u8>, JsError> {
    let js_edits: Vec<JsSlideTextEdit> = serde_wasm_bindgen::from_value(edits)
        .map_err(|e| JsError::new(&e.to_string()))?;

    let mut editor = oxislides_core::PptxEditor::new(data)
        .map_err(|e| JsError::new(&e.to_string()))?;

    let edits: Vec<oxislides_core::editor::SlideTextEdit> = js_edits
        .into_iter()
        .map(|e| oxislides_core::editor::SlideTextEdit {
            slide_index: e.slide_index,
            shape_index: e.shape_index,
            paragraph_index: e.paragraph_index,
            run_index: e.run_index,
            new_text: e.new_text,
        })
        .collect();

    editor.apply_edits(&edits);

    editor
        .save()
        .map_err(|e| JsError::new(&e.to_string()))
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
    font_family: Option<String>,
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
                            text, font_size, font_family, bold, italic, underline, strikethrough, color, highlight,
                        } => LayoutElementJs {
                            x: elem.x, y: elem.y, width: elem.width, height: elem.height,
                            kind: "text".into(),
                            text: Some(text),
                            font_size: Some(font_size),
                            font_family,
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
                                text: None, font_size: None, font_family: None, bold: None, italic: None,
                                underline: None, strikethrough: None, color: None, highlight: None,
                                image_data: b64,
                                content_type,
                                x1: None, y1: None, x2: None, y2: None,
                            }
                        },
                        oxidocs_core::layout::LayoutContent::TableBorder { x1, y1, x2, y2 } => LayoutElementJs {
                            x: elem.x, y: elem.y, width: elem.width, height: elem.height,
                            kind: "border".into(),
                            text: None, font_size: None, font_family: None, bold: None, italic: None,
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

// ---------------------------------------------------------------------------
// PDF bindings (oxipdf-core)
// ---------------------------------------------------------------------------

/// Parse a PDF file and return its structure as a JS object.
#[wasm_bindgen]
pub fn parse_pdf(data: &[u8]) -> Result<JsValue, JsError> {
    let doc = oxipdf_core::parse_pdf(data)
        .map_err(|e| JsError::new(&e.to_string()))?;
    serde_wasm_bindgen::to_value(&doc).map_err(|e| JsError::new(&e.to_string()))
}

/// Extract all text from a PDF as a single string.
#[wasm_bindgen]
pub fn pdf_extract_text(data: &[u8]) -> Result<String, JsError> {
    let doc = oxipdf_core::parse_pdf(data)
        .map_err(|e| JsError::new(&e.to_string()))?;
    Ok(oxipdf_core::extract_text_string(&doc))
}

/// Generate a PDF from scratch with the given text content.
/// Returns the PDF bytes.
#[wasm_bindgen]
pub fn create_pdf(title: &str, text: &str) -> Vec<u8> {
    use oxipdf_core::ir::*;

    let doc = PdfDocument {
        version: PdfVersion::new(1, 7),
        info: DocumentInfo {
            title: Some(title.to_string()),
            producer: Some("Oxi".to_string()),
            ..Default::default()
        },
        pages: vec![Page {
            width: 595.0,  // A4
            height: 842.0,
            media_box: Rectangle { llx: 0.0, lly: 0.0, urx: 595.0, ury: 842.0 },
            crop_box: None,
            contents: vec![ContentElement::Text(TextSpan {
                x: 72.0,
                y: 72.0,
                text: text.to_string(),
                font_name: "Helvetica".to_string(),
                font_size: 12.0,
                fill_color: Color::Gray(0.0),
            })],
            rotation: 0,
        }],
        outline: Vec::new(),
    };
    oxipdf_core::write_pdf(&doc)
}

/// Verify signatures in a PDF. Returns an array of signature info objects.
#[wasm_bindgen]
pub fn pdf_verify_signatures(data: &[u8]) -> Result<JsValue, JsError> {
    let sigs = oxipdf_core::verify_pdf_signatures(data)
        .map_err(|e| JsError::new(&e.to_string()))?;
    serde_wasm_bindgen::to_value(&sigs).map_err(|e| JsError::new(&e.to_string()))
}

// ---------------------------------------------------------------------------
// Hanko bindings (oxihanko)
// ---------------------------------------------------------------------------

/// Generate a hanko stamp SVG.
///
/// `config`: JS object with StampConfig fields:
///   { name: "山田", style: "Round"|"Square"|"Oval", size: 100, date?: "2026.03.13" }
#[wasm_bindgen]
pub fn generate_hanko_svg(config: JsValue) -> Result<String, JsError> {
    let stamp_config: oxihanko::StampConfig = serde_wasm_bindgen::from_value(config)
        .map_err(|e| JsError::new(&e.to_string()))?;
    Ok(oxihanko::generate_stamp_svg(&stamp_config))
}

/// Preview a hanko stamp SVG with default config for the given name.
#[wasm_bindgen]
pub fn preview_hanko(name: &str) -> String {
    let config = oxihanko::StampConfig {
        name: name.to_string(),
        ..Default::default()
    };
    oxihanko::generate_stamp_svg(&config)
}
