use wasm_bindgen::prelude::*;
use serde::{Deserialize, Serialize};
use std::sync::OnceLock;

static OXI_GOTHIC: &[u8] = include_bytes!("../../oxi-cli/fonts/OxiGothic.ttf");
static OXI_MINCHO: &[u8] = include_bytes!("../../oxi-cli/fonts/OxiMincho.ttf");

fn get_cjk_fonts() -> &'static (oxipdf_core::ir::EmbeddedFont, oxipdf_core::ir::EmbeddedFont) {
    static FONTS: OnceLock<(oxipdf_core::ir::EmbeddedFont, oxipdf_core::ir::EmbeddedFont)> = OnceLock::new();
    FONTS.get_or_init(|| {
        let gothic = oxipdf_core::font_util::embedded_font_from_ttf(OXI_GOTHIC);
        let mincho = oxipdf_core::font_util::embedded_font_from_ttf(OXI_MINCHO);
        (gothic, mincho)
    })
}

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

/// Build a .docx from a content structure.
/// `content`: JS array of block objects:
///   { type: "paragraph", runs: [{text, bold?, italic?, underline?, strikethrough?, font_family?, font_size?, color?}], alignment?, heading_level?, line_height? }
///   { type: "table", rows: [[{text, bold?}]] }
#[wasm_bindgen]
pub fn build_docx(content: JsValue) -> Result<Vec<u8>, JsError> {
    let blocks: Vec<oxidocs_core::ContentBlock> = serde_wasm_bindgen::from_value(content)
        .map_err(|e| JsError::new(&e.to_string()))?;
    Ok(oxidocs_core::build_docx(&blocks))
}

/// Build a .docx from content, using a template docx for styles/theme/numbering.
/// Preserves original formatting while replacing document content.
#[wasm_bindgen]
pub fn build_docx_with_template(content: JsValue, template: &[u8]) -> Result<Vec<u8>, JsError> {
    let blocks: Vec<oxidocs_core::ContentBlock> = serde_wasm_bindgen::from_value(content)
        .map_err(|e| JsError::new(&e.to_string()))?;
    Ok(oxidocs_core::build_docx_with_template(&blocks, template))
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

// ---------------------------------------------------------------------------
// Advanced docx editing bindings
// ---------------------------------------------------------------------------

/// Apply structural edits to a .docx file.
///
/// `data`: original .docx bytes
/// `edits`: JS array of edit operation objects. Each object has a `type` field:
///
/// Text operations:
///   { type: "set_run_text", paragraph_index, run_index, new_text }
///   { type: "insert_paragraph", index, text, style?, para_style? }
///   { type: "delete_paragraph", index }
///   { type: "insert_run", paragraph_index, run_index, text, style? }
///   { type: "delete_run", paragraph_index, run_index }
///
/// Formatting:
///   { type: "set_run_format", paragraph_index, run_index, style }
///   { type: "set_paragraph_format", paragraph_index, style }
///
/// Tables:
///   { type: "insert_table", index, rows, cols, content?, col_widths_pt? }
///   { type: "insert_table_row", table_index, row_index, cells }
///   { type: "delete_table_row", table_index, row_index }
///   { type: "set_cell_text", table_index, row, col, text }
///
/// Images:
///   { type: "insert_image", index, data (base64), width_pt, height_pt, content_type }
///
/// style (RunProps): { bold?, italic?, underline?, font_family?, font_size?, color?, highlight? }
/// para_style (ParaProps): { alignment?, space_before?, space_after?, line_spacing?, indent_left?, style_id? }
#[wasm_bindgen]
pub fn edit_docx_advanced(data: &[u8], edits: JsValue) -> Result<Vec<u8>, JsError> {
    let js_edits: Vec<JsDocxEdit> = serde_wasm_bindgen::from_value(edits)
        .map_err(|e| JsError::new(&e.to_string()))?;

    let mut editor = oxidocs_core::DocxEditor::new(data)
        .map_err(|e| JsError::new(&e.to_string()))?;

    for js_edit in js_edits {
        match js_edit.r#type.as_str() {
            "set_run_text" => {
                editor.add_edit(oxidocs_core::editor::DocxEdit::SetRunText {
                    paragraph_index: js_edit.paragraph_index.unwrap_or(0),
                    run_index: js_edit.run_index.unwrap_or(0),
                    new_text: js_edit.new_text.unwrap_or_default(),
                });
            }
            "insert_paragraph" => {
                let run_style = js_edit.style.map(|s| to_run_props(&s));
                let para_style = js_edit.para_style.map(|s| to_para_props(&s));
                let text = js_edit.text.unwrap_or_default();
                editor.add_edit(oxidocs_core::editor::DocxEdit::InsertParagraph {
                    index: js_edit.index.unwrap_or(0),
                    runs: vec![(text, run_style)],
                    style: para_style,
                });
            }
            "delete_paragraph" => {
                editor.add_edit(oxidocs_core::editor::DocxEdit::DeleteParagraph {
                    index: js_edit.index.unwrap_or(0),
                });
            }
            "insert_run" => {
                let style = js_edit.style.map(|s| to_run_props(&s));
                editor.add_edit(oxidocs_core::editor::DocxEdit::InsertRun {
                    paragraph_index: js_edit.paragraph_index.unwrap_or(0),
                    run_index: js_edit.run_index.unwrap_or(0),
                    text: js_edit.text.unwrap_or_default(),
                    style,
                });
            }
            "delete_run" => {
                editor.add_edit(oxidocs_core::editor::DocxEdit::DeleteRun {
                    paragraph_index: js_edit.paragraph_index.unwrap_or(0),
                    run_index: js_edit.run_index.unwrap_or(0),
                });
            }
            "set_run_format" => {
                let style = js_edit.style.map(|s| to_run_props(&s)).unwrap_or_default();
                editor.add_edit(oxidocs_core::editor::DocxEdit::SetRunFormat {
                    paragraph_index: js_edit.paragraph_index.unwrap_or(0),
                    run_index: js_edit.run_index.unwrap_or(0),
                    style,
                });
            }
            "set_paragraph_format" => {
                let style = js_edit.para_style.map(|s| to_para_props(&s)).unwrap_or_default();
                editor.add_edit(oxidocs_core::editor::DocxEdit::SetParagraphFormat {
                    paragraph_index: js_edit.paragraph_index.unwrap_or(0),
                    style,
                });
            }
            "insert_table" => {
                editor.add_edit(oxidocs_core::editor::DocxEdit::InsertTable {
                    index: js_edit.index.unwrap_or(0),
                    rows: js_edit.rows.unwrap_or(1),
                    cols: js_edit.cols.unwrap_or(1),
                    content: js_edit.content,
                    col_widths_pt: js_edit.col_widths_pt,
                });
            }
            "insert_table_row" => {
                editor.add_edit(oxidocs_core::editor::DocxEdit::InsertTableRow {
                    table_index: js_edit.table_index.unwrap_or(0),
                    row_index: js_edit.row_index.unwrap_or(0),
                    cells: js_edit.cells.unwrap_or_default(),
                });
            }
            "delete_table_row" => {
                editor.add_edit(oxidocs_core::editor::DocxEdit::DeleteTableRow {
                    table_index: js_edit.table_index.unwrap_or(0),
                    row_index: js_edit.row_index.unwrap_or(0),
                });
            }
            "set_cell_text" => {
                editor.add_edit(oxidocs_core::editor::DocxEdit::SetCellText {
                    table_index: js_edit.table_index.unwrap_or(0),
                    row: js_edit.row.unwrap_or(0),
                    col: js_edit.col.unwrap_or(0),
                    text: js_edit.text.unwrap_or_default(),
                });
            }
            "insert_image" => {
                let image_data = js_edit.image_data.map(|b64| base64_decode(&b64)).unwrap_or_default();
                let ct = js_edit.content_type_field.unwrap_or_else(|| "image/png".to_string());
                editor.insert_image(
                    js_edit.index.unwrap_or(0),
                    image_data,
                    js_edit.width_pt.unwrap_or(100.0),
                    js_edit.height_pt.unwrap_or(100.0),
                    &ct,
                );
            }
            other => {
                return Err(JsError::new(&format!("Unknown edit type: {}", other)));
            }
        }
    }

    editor.save().map_err(|e| JsError::new(&e.to_string()))
}

#[derive(Deserialize)]
struct JsDocxEdit {
    r#type: String,
    // Common fields
    #[serde(default)]
    index: Option<usize>,
    #[serde(default)]
    paragraph_index: Option<usize>,
    #[serde(default)]
    run_index: Option<usize>,
    #[serde(default)]
    new_text: Option<String>,
    #[serde(default)]
    text: Option<String>,
    // Formatting
    #[serde(default)]
    style: Option<JsRunProps>,
    #[serde(default)]
    para_style: Option<JsParaProps>,
    // Table
    #[serde(default)]
    table_index: Option<usize>,
    #[serde(default)]
    row_index: Option<usize>,
    #[serde(default)]
    rows: Option<usize>,
    #[serde(default)]
    cols: Option<usize>,
    #[serde(default)]
    row: Option<usize>,
    #[serde(default)]
    col: Option<usize>,
    #[serde(default)]
    content: Option<Vec<Vec<String>>>,
    #[serde(default)]
    cells: Option<Vec<String>>,
    #[serde(default)]
    col_widths_pt: Option<Vec<f32>>,
    // Image
    #[serde(default, rename = "data")]
    image_data: Option<String>,
    #[serde(default)]
    width_pt: Option<f32>,
    #[serde(default)]
    height_pt: Option<f32>,
    #[serde(default, rename = "content_type")]
    content_type_field: Option<String>,
}

#[derive(Deserialize)]
struct JsRunProps {
    #[serde(default)]
    bold: Option<bool>,
    #[serde(default)]
    italic: Option<bool>,
    #[serde(default)]
    underline: Option<bool>,
    #[serde(default)]
    underline_style: Option<String>,
    #[serde(default)]
    strikethrough: Option<bool>,
    #[serde(default)]
    font_family: Option<String>,
    #[serde(default)]
    font_family_east_asia: Option<String>,
    #[serde(default)]
    font_size: Option<f32>,
    #[serde(default)]
    color: Option<String>,
    #[serde(default)]
    highlight: Option<String>,
    #[serde(default)]
    character_spacing: Option<f32>,
    #[serde(default)]
    kerning: Option<f32>,
    #[serde(default)]
    lang: Option<String>,
    #[serde(default)]
    lang_east_asia: Option<String>,
    #[serde(default)]
    lang_bidi: Option<String>,
    #[serde(default)]
    run_style: Option<String>,
    #[serde(default)]
    no_proof: Option<bool>,
    #[serde(default)]
    vertical_align: Option<String>,
}

#[derive(Deserialize)]
struct JsParaProps {
    #[serde(default)]
    alignment: Option<String>,
    #[serde(default)]
    space_before: Option<f32>,
    #[serde(default)]
    space_after: Option<f32>,
    #[serde(default)]
    line_spacing: Option<f32>,
    #[serde(default)]
    indent_left: Option<f32>,
    #[serde(default)]
    indent_right: Option<f32>,
    #[serde(default)]
    indent_first_line: Option<f32>,
    #[serde(default)]
    style_id: Option<String>,
    #[serde(default)]
    keep_next: Option<bool>,
    #[serde(default)]
    keep_lines: Option<bool>,
    #[serde(default)]
    widow_control: Option<bool>,
    #[serde(default)]
    snap_to_grid: Option<bool>,
    #[serde(default)]
    word_wrap: Option<bool>,
    #[serde(default)]
    adjust_right_ind: Option<bool>,
    #[serde(default)]
    auto_space_de: Option<bool>,
    #[serde(default)]
    auto_space_dn: Option<bool>,
    #[serde(default)]
    page_break_before: Option<bool>,
}

fn to_run_props(js: &JsRunProps) -> oxidocs_core::editor::RunProps {
    oxidocs_core::editor::RunProps {
        bold: js.bold,
        italic: js.italic,
        underline: js.underline,
        underline_style: js.underline_style.clone(),
        strikethrough: js.strikethrough,
        font_family: js.font_family.clone(),
        font_family_east_asia: js.font_family_east_asia.clone(),
        font_size: js.font_size,
        color: js.color.clone(),
        highlight: js.highlight.clone(),
        character_spacing: js.character_spacing,
        kerning: js.kerning,
        lang: js.lang.clone(),
        lang_east_asia: js.lang_east_asia.clone(),
        lang_bidi: js.lang_bidi.clone(),
        run_style: js.run_style.clone(),
        no_proof: js.no_proof,
        vertical_align: js.vertical_align.clone(),
    }
}

fn to_para_props(js: &JsParaProps) -> oxidocs_core::editor::ParaProps {
    oxidocs_core::editor::ParaProps {
        alignment: js.alignment.clone(),
        space_before: js.space_before,
        space_after: js.space_after,
        line_spacing: js.line_spacing,
        indent_left: js.indent_left,
        indent_right: js.indent_right,
        indent_first_line: js.indent_first_line,
        style_id: js.style_id.clone(),
        keep_next: js.keep_next,
        keep_lines: js.keep_lines,
        widow_control: js.widow_control,
        snap_to_grid: js.snap_to_grid,
        word_wrap: js.word_wrap,
        adjust_right_ind: js.adjust_right_ind,
        auto_space_de: js.auto_space_de,
        auto_space_dn: js.auto_space_dn,
        page_break_before: js.page_break_before,
    }
}

fn base64_decode(input: &str) -> Vec<u8> {
    let mut out = Vec::with_capacity(input.len() * 3 / 4);
    let mut buf: u32 = 0;
    let mut bits: u32 = 0;
    for &b in input.as_bytes() {
        let val = match b {
            b'A'..=b'Z' => b - b'A',
            b'a'..=b'z' => b - b'a' + 26,
            b'0'..=b'9' => b - b'0' + 52,
            b'+' => 62,
            b'/' => 63,
            b'=' | b'\n' | b'\r' | b' ' => continue,
            _ => continue,
        };
        buf = (buf << 6) | val as u32;
        bits += 6;
        if bits >= 8 {
            bits -= 8;
            out.push((buf >> bits) as u8);
            buf &= (1 << bits) - 1;
        }
    }
    out
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
    underline_style: Option<String>,
    strikethrough: Option<bool>,
    color: Option<String>,
    highlight: Option<String>,
    character_spacing: Option<f32>,
    // Box fields
    corner_radius: Option<f32>,
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

    let engine = oxidocs_core::layout::LayoutEngine::for_document(&doc);
    let result = engine.layout(&doc);

    let js_result = LayoutResultJs {
        pages: result.pages.into_iter().map(|page| {
            LayoutPageJs {
                width: page.width,
                height: page.height,
                elements: page.elements.into_iter().map(|elem| {
                    match elem.content {
                        oxidocs_core::layout::LayoutContent::Text {
                            text, font_size, font_family, bold, italic, underline, underline_style, strikethrough, color, highlight, character_spacing, ..
                        } => LayoutElementJs {
                            x: elem.x, y: elem.y, width: elem.width, height: elem.height,
                            kind: "text".into(),
                            text: Some(text),
                            font_size: Some(font_size),
                            font_family,
                            bold: Some(bold),
                            italic: Some(italic),
                            underline: Some(underline),
                            underline_style,
                            strikethrough: Some(strikethrough),
                            color: color.map(|c| if c.starts_with('#') { c } else { format!("#{}", c) }),
                            highlight,
                            character_spacing: if character_spacing.abs() > 0.001 { Some(character_spacing) } else { None },
                            corner_radius: None,
                            image_data: None, content_type: None,
                            x1: None, y1: None, x2: None, y2: None,
                        },
                        oxidocs_core::layout::LayoutContent::Image { data, content_type } => {
                            let b64 = if !data.is_empty() { Some(base64_encode(&data)) } else { None };
                            LayoutElementJs {
                                x: elem.x, y: elem.y, width: elem.width, height: elem.height,
                                kind: "image".into(),
                                text: None, font_size: None, font_family: None, bold: None, italic: None,
                                underline: None, underline_style: None, strikethrough: None, color: None, highlight: None,
                                character_spacing: None, corner_radius: None,
                                image_data: b64,
                                content_type,
                                x1: None, y1: None, x2: None, y2: None,
                            }
                        },
                        oxidocs_core::layout::LayoutContent::TableBorder { x1, y1, x2, y2, color, width } => LayoutElementJs {
                            x: elem.x, y: elem.y, width: width, height: elem.height,
                            kind: "border".into(),
                            text: None, font_size: None, font_family: None, bold: None, italic: None,
                            underline: None, underline_style: None, strikethrough: None, color, highlight: None,
                            character_spacing: None, corner_radius: None,
                            image_data: None, content_type: None,
                            x1: Some(x1), y1: Some(y1), x2: Some(x2), y2: Some(y2),
                        },
                        oxidocs_core::layout::LayoutContent::CellShading { color } => LayoutElementJs {
                            x: elem.x, y: elem.y, width: elem.width, height: elem.height,
                            kind: "shading".into(),
                            text: None, font_size: None, font_family: None, bold: None, italic: None,
                            underline: None, underline_style: None, strikethrough: None, color: Some(color), highlight: None,
                            character_spacing: None, corner_radius: None,
                            image_data: None, content_type: None,
                            x1: None, y1: None, x2: None, y2: None,
                        },
                        oxidocs_core::layout::LayoutContent::BoxRect { fill, stroke_color, corner_radius, .. } => LayoutElementJs {
                            x: elem.x, y: elem.y, width: elem.width, height: elem.height,
                            kind: "shading".into(),
                            text: None, font_size: None, font_family: None, bold: None, italic: None,
                            underline: None, underline_style: None, strikethrough: None,
                            color: fill.clone().or_else(|| stroke_color.clone()),
                            highlight: None,
                            character_spacing: None, corner_radius: if corner_radius > 0.0 { Some(corner_radius) } else { None },
                            image_data: None, content_type: None,
                            x1: None, y1: None, x2: None, y2: None,
                        },
                        oxidocs_core::layout::LayoutContent::ClipStart => LayoutElementJs {
                            x: elem.x, y: elem.y, width: elem.width, height: elem.height,
                            kind: "clip_start".into(),
                            text: None, font_size: None, font_family: None, bold: None, italic: None,
                            underline: None, underline_style: None, strikethrough: None,
                            color: None, highlight: None, character_spacing: None, corner_radius: None,
                            image_data: None, content_type: None,
                            x1: None, y1: None, x2: None, y2: None,
                        },
                        oxidocs_core::layout::LayoutContent::ClipEnd => LayoutElementJs {
                            x: 0.0, y: 0.0, width: 0.0, height: 0.0,
                            kind: "clip_end".into(),
                            text: None, font_size: None, font_family: None, bold: None, italic: None,
                            underline: None, underline_style: None, strikethrough: None,
                            color: None, highlight: None, character_spacing: None, corner_radius: None,
                            image_data: None, content_type: None,
                            x1: None, y1: None, x2: None, y2: None,
                        },
                        oxidocs_core::layout::LayoutContent::PresetShape { .. } => LayoutElementJs {
                            x: elem.x, y: elem.y, width: elem.width, height: elem.height,
                            kind: "preset_shape".into(),
                            text: None, font_size: None, font_family: None, bold: None, italic: None,
                            underline: None, underline_style: None, strikethrough: None,
                            color: None, highlight: None, character_spacing: None, corner_radius: None,
                            image_data: None, content_type: None,
                            x1: None, y1: None, x2: None, y2: None,
                        },
                    }
                }).collect(),
            }
        }).collect(),
    };

    serde_wasm_bindgen::to_value(&js_result).map_err(|e| JsError::new(&e.to_string()))
}

// ---------------------------------------------------------------------------
// docx → PDF conversion
// ---------------------------------------------------------------------------

/// Convert a .docx file to PDF bytes.
/// Parses the docx, runs layout, and converts positioned elements to PDF.
#[wasm_bindgen]
pub fn docx_to_pdf(data: &[u8]) -> Result<Vec<u8>, JsError> {
    let doc = oxidocs_core::parse_docx(data)
        .map_err(|e| JsError::new(&e.to_string()))?;

    let engine = oxidocs_core::layout::LayoutEngine::for_document(&doc);
    let layout = engine.layout(&doc);

    let pdf_doc = layout_to_pdf(&layout, &doc);
    Ok(oxipdf_core::write_pdf(&pdf_doc))
}

fn layout_to_pdf(
    layout: &oxidocs_core::layout::LayoutResult,
    doc: &oxidocs_core::Document,
) -> oxipdf_core::ir::PdfDocument {
    use oxipdf_core::ir::*;

    let title = doc.pages.first()
        .and_then(|p| p.blocks.first())
        .and_then(|b| match b {
            oxidocs_core::ir::Block::Paragraph(p) if p.style.heading_level == Some(1) => {
                Some(p.runs.iter().map(|r| r.text.as_str()).collect::<String>())
            }
            _ => None,
        });

    let pages = layout.pages.iter().map(|lp| {
        let width = lp.width as f64;
        let height = lp.height as f64;
        let mut contents = Vec::new();

        for elem in &lp.elements {
            match &elem.content {
                oxidocs_core::layout::LayoutContent::Text {
                    text, font_size, font_family, bold, color, character_spacing, ..
                } => {
                    if text.is_empty() {
                        continue;
                    }
                    let base_font = font_family.as_deref().unwrap_or("Helvetica");
                    let font_name = resolve_wasm_font(base_font, *bold, text);

                    let fill_color = color.as_ref()
                        .and_then(|c| parse_hex_color(c))
                        .unwrap_or(Color::Gray(0.0));

                    contents.push(ContentElement::Text(TextSpan {
                        x: elem.x as f64,
                        y: elem.y as f64,
                        text: text.clone(),
                        font_name,
                        font_size: *font_size as f64,
                        fill_color,
                        character_spacing: *character_spacing as f64,
                    }));
                }
                oxidocs_core::layout::LayoutContent::Image { data, .. } => {
                    if data.is_empty() {
                        continue;
                    }
                    contents.push(ContentElement::Image(ImageData {
                        x: elem.x as f64,
                        y: elem.y as f64,
                        width: elem.width as f64,
                        height: elem.height as f64,
                        data: data.clone(),
                        color_space: ColorSpace::DeviceRgb,
                        bits_per_component: 8,
                        pixel_width: elem.width as u32,
                        pixel_height: elem.height as u32,
                    }));
                }
                oxidocs_core::layout::LayoutContent::TableBorder { x1, y1, x2, y2, ref color, width } => {
                    let stroke_color = color.as_ref()
                        .and_then(|c| parse_hex_color(c))
                        .unwrap_or(Color::Gray(0.0));
                    contents.push(ContentElement::Path(PathData {
                        operations: vec![
                            PathOp::MoveTo(*x1 as f64, *y1 as f64),
                            PathOp::LineTo(*x2 as f64, *y2 as f64),
                        ],
                        stroke: Some(StrokeStyle {
                            color: stroke_color,
                            width: *width as f64,
                            line_cap: LineCap::Butt,
                            line_join: LineJoin::Miter,
                        }),
                        fill: None,
                    }));
                }
                oxidocs_core::layout::LayoutContent::CellShading { ref color } => {
                    if let Some(fill_color) = parse_hex_color(color) {
                        contents.push(ContentElement::Path(PathData {
                            operations: vec![
                                PathOp::MoveTo(elem.x as f64, elem.y as f64),
                                PathOp::LineTo((elem.x + elem.width) as f64, elem.y as f64),
                                PathOp::LineTo((elem.x + elem.width) as f64, (elem.y + elem.height) as f64),
                                PathOp::LineTo(elem.x as f64, (elem.y + elem.height) as f64),
                                PathOp::ClosePath,
                            ],
                            stroke: None,
                            fill: Some(fill_color),
                        }));
                    }
                }
                oxidocs_core::layout::LayoutContent::BoxRect { ref fill, ref stroke_color, stroke_width, .. } => {
                    let fill_color = fill.as_deref().and_then(parse_hex_color);
                    let stroke = stroke_color.as_deref().and_then(parse_hex_color).map(|c| StrokeStyle {
                        color: c, width: *stroke_width as f64,
                        line_cap: LineCap::Butt, line_join: LineJoin::Miter,
                    });
                    contents.push(ContentElement::Path(PathData {
                        operations: vec![
                            PathOp::MoveTo(elem.x as f64, elem.y as f64),
                            PathOp::LineTo((elem.x + elem.width) as f64, elem.y as f64),
                            PathOp::LineTo((elem.x + elem.width) as f64, (elem.y + elem.height) as f64),
                            PathOp::LineTo(elem.x as f64, (elem.y + elem.height) as f64),
                            PathOp::ClosePath,
                        ],
                        stroke,
                        fill: fill_color,
                    }));
                }
                oxidocs_core::layout::LayoutContent::ClipStart | oxidocs_core::layout::LayoutContent::ClipEnd => {}
                oxidocs_core::layout::LayoutContent::PresetShape { .. } => {}
            }
        }

        Page {
            width,
            height,
            media_box: Rectangle { llx: 0.0, lly: 0.0, urx: width, ury: height },
            crop_box: None,
            contents,
            rotation: 0,
        }
    }).collect::<Vec<Page>>();

    // Check if any text span uses CJK fonts
    let has_gothic = pages.iter().any(|p| p.contents.iter().any(|c| {
        matches!(c, ContentElement::Text(t) if t.font_name == "OxiCJK-Gothic")
    }));
    let has_mincho = pages.iter().any(|p| p.contents.iter().any(|c| {
        matches!(c, ContentElement::Text(t) if t.font_name == "OxiCJK-Mincho")
    }));

    let mut embedded_fonts = std::collections::HashMap::new();
    if has_gothic || has_mincho {
        let (gothic, mincho) = get_cjk_fonts();
        if has_gothic {
            embedded_fonts.insert("OxiCJK-Gothic".to_string(), gothic.clone());
        }
        if has_mincho {
            embedded_fonts.insert("OxiCJK-Mincho".to_string(), mincho.clone());
        }
    }

    PdfDocument {
        version: PdfVersion::new(1, 7),
        info: DocumentInfo {
            title,
            producer: Some("Oxi (docx→PDF)".to_string()),
            ..Default::default()
        },
        pages,
        outline: Vec::new(),
        embedded_fonts,
    }
}

fn resolve_wasm_font(base_font: &str, bold: bool, text: &str) -> String {
    // Check if font is a known CJK family
    let is_cjk_font = matches!(base_font,
        "ＭＳ ゴシック" | "MS ゴシック" | "MS Gothic" | "ＭＳ Ｐゴシック" | "MS PGothic"
        | "Yu Gothic" | "游ゴシック" | "メイリオ" | "Meiryo" | "ヒラギノ角ゴ"
        | "ＭＳ 明朝" | "MS 明朝" | "MS Mincho" | "ＭＳ Ｐ明朝" | "MS PMincho"
        | "Yu Mincho" | "游明朝" | "ヒラギノ明朝"
    );

    if is_cjk_font {
        if base_font.contains("明朝") || base_font.contains("Mincho") {
            return "OxiCJK-Mincho".to_string();
        }
        return "OxiCJK-Gothic".to_string();
    }

    // Check if text contains CJK characters (fallback)
    let has_cjk = text.chars().any(|c| c as u32 > 0x2E7F);
    if has_cjk {
        return "OxiCJK-Gothic".to_string();
    }

    // Standard PDF fonts
    if bold {
        match base_font {
            "Helvetica" | "Calibri" | "Arial" => "Helvetica-Bold".to_string(),
            "Times New Roman" => "Times-Bold".to_string(),
            "Courier New" | "Courier" => "Courier-Bold".to_string(),
            _ => base_font.to_string(),
        }
    } else {
        match base_font {
            "Times New Roman" => "Times-Roman".to_string(),
            "Courier New" => "Courier".to_string(),
            _ => base_font.to_string(),
        }
    }
}

/// Parse "#RRGGBB" hex color to PDF Color.
fn parse_hex_color(hex: &str) -> Option<oxipdf_core::ir::Color> {
    let hex = hex.strip_prefix('#').unwrap_or(hex);
    if hex.len() != 6 {
        return None;
    }
    let r = u8::from_str_radix(&hex[0..2], 16).ok()? as f64 / 255.0;
    let g = u8::from_str_radix(&hex[2..4], 16).ok()? as f64 / 255.0;
    let b = u8::from_str_radix(&hex[4..6], 16).ok()? as f64 / 255.0;
    Some(oxipdf_core::ir::Color::Rgb(r, g, b))
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
                fill_color: Color::Gray(0.0), character_spacing: 0.0,
            })],
            rotation: 0,
        }],
        outline: Vec::new(),
        embedded_fonts: std::collections::HashMap::new(),
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
