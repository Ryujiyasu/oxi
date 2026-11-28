// SPDX-License-Identifier: MIT OR Apache-2.0

//! Spreadsheet print pages as PDF in the browser; the pages themselves come
//! from `oxidocs_render::xlsx`, shared with the native viewers.

use wasm_bindgen::prelude::*;

/// Convert a workbook (.xlsx, .xlsm, .xltx, .xltm) to PDF bytes: the print
/// areas of each visible sheet, or its used range, laid out with the sheet's
/// page setup from the values saved in the file. Formulas are not
/// recalculated. `title` becomes the PDF's title when it is not empty.
#[wasm_bindgen]
pub fn xlsx_to_pdf(data: &[u8], title: &str) -> Result<Vec<u8>, JsError> {
    let doc = oxidocs_render::xlsx::workbook_pages(data, title).map_err(|e| JsError::new(&e))?;
    oxipdf_core::write_pdf_checked(&doc).map_err(|e| JsError::new(&format!("PDF glyph encoding: {e:?}")))
}
