//! Mail merge for the page: read the records (CSV/TSV text or an Excel
//! workbook) and merge a main document with them.

use oxidocs_core::mail_merge::{self, MergeData, MergeOptions};
use wasm_bindgen::prelude::*;

fn js_err<E: std::fmt::Display>(e: E) -> JsError {
    JsError::new(&e.to_string())
}

fn options(opts: JsValue) -> Result<MergeOptions, JsError> {
    if opts.is_undefined() || opts.is_null() {
        Ok(MergeOptions::default())
    } else {
        serde_wasm_bindgen::from_value(opts).map_err(js_err)
    }
}

/// The merge field names a main document uses: `{ fields: [...] }`.
#[wasm_bindgen]
pub fn mail_merge_fields(docx: &[u8]) -> Result<JsValue, JsError> {
    let report = mail_merge::merge_fields(docx).map_err(js_err)?;
    serde_wasm_bindgen::to_value(&report).map_err(js_err)
}

/// Records from comma- or tab-separated text: `{ columns, rows }`.
#[wasm_bindgen]
pub fn mail_merge_data_from_text(text: &str) -> Result<JsValue, JsError> {
    serde_wasm_bindgen::to_value(&mail_merge::parse_delimited(text)).map_err(js_err)
}

/// Merge every record into one document (Word's "merge to new document").
/// `data` is `{ columns, rows }`; `opts` is `{ suppress_blank_lines,
/// first_record, last_record }` (all optional).
#[wasm_bindgen]
pub fn mail_merge_docx(docx: &[u8], data: JsValue, opts: JsValue) -> Result<Vec<u8>, JsError> {
    let data: MergeData = serde_wasm_bindgen::from_value(data).map_err(js_err)?;
    mail_merge::mail_merge(docx, &data, &options(opts)?).map_err(js_err)
}

/// Merge each record into a document of its own, all of them in one .zip
/// named `<stem>_<record>.docx`.
#[cfg(feature = "suite")]
#[wasm_bindgen]
pub fn mail_merge_docx_each(docx: &[u8], data: JsValue, opts: JsValue, stem: &str) -> Result<Vec<u8>, JsError> {
    use std::io::Write;
    let data: MergeData = serde_wasm_bindgen::from_value(data).map_err(js_err)?;
    let docs = mail_merge::mail_merge_each(docx, &data, &options(opts)?).map_err(js_err)?;
    let width = data.rows.len().max(1).to_string().len();
    let mut zip = zip::ZipWriter::new(std::io::Cursor::new(Vec::new()));
    let o = zip::write::SimpleFileOptions::default().compression_method(zip::CompressionMethod::Stored);
    for d in docs {
        zip.start_file(format!("{stem}_{:0width$}.docx", d.record), o).map_err(js_err)?;
        zip.write_all(&d.bytes).map_err(js_err)?;
    }
    Ok(zip.finish().map_err(js_err)?.into_inner())
}

/// Merge one record (counting from 1) into a document, to look at.
#[wasm_bindgen]
pub fn mail_merge_docx_record(docx: &[u8], data: JsValue, record: usize, opts: JsValue) -> Result<Vec<u8>, JsError> {
    let data: MergeData = serde_wasm_bindgen::from_value(data).map_err(js_err)?;
    let mut o = options(opts)?;
    o.first_record = Some(record);
    o.last_record = Some(data.rows.len());
    let mut docs = mail_merge::mail_merge_each(docx, &data, &o).map_err(js_err)?;
    if docs.is_empty() {
        return Err(JsError::new("no record to merge"));
    }
    Ok(docs.swap_remove(0).bytes)
}

/// The field names the main document asks for that the data lacks.
#[wasm_bindgen]
pub fn mail_merge_missing_fields(docx: &[u8], data: JsValue) -> Result<JsValue, JsError> {
    let data: MergeData = serde_wasm_bindgen::from_value(data).map_err(js_err)?;
    let missing = mail_merge::missing_fields(docx, &data).map_err(js_err)?;
    serde_wasm_bindgen::to_value(&missing).map_err(js_err)
}

/// Records from a sheet of an Excel workbook: `{ sheets, sheet, columns,
/// rows }`. The first row names the columns. Values come as Word reads them
/// from Excel (measured on Word's own merges): the stored value, not the
/// displayed one -- a number in its plain form (`1234.5` for `1,235`, `0.25`
/// for `25%`), a date as `m/d/yyyy`, a time as `h:mm:ss AM/PM`, a boolean
/// as `True`/`False`, a formula as its value.
#[cfg(feature = "suite")]
#[wasm_bindgen]
pub fn mail_merge_data_from_xlsx(xlsx: &[u8], sheet: Option<usize>) -> Result<JsValue, JsError> {
    #[derive(serde::Serialize)]
    struct Out {
        sheets: Vec<String>,
        sheet: usize,
        columns: Vec<String>,
        rows: Vec<Vec<String>>,
    }
    let wb = oxicells_core::parser::parse_xlsx_preserving_values(xlsx).map_err(js_err)?;
    let sheets: Vec<String> = wb.sheets.iter().map(|s| s.name.clone()).collect();
    let index = sheet.unwrap_or(0).min(sheets.len().saturating_sub(1));
    let data = wb.sheets.get(index).map(sheet_records).unwrap_or_default();
    serde_wasm_bindgen::to_value(&Out { sheets, sheet: index, columns: data.columns, rows: data.rows }).map_err(js_err)
}

#[cfg(feature = "suite")]
fn sheet_records(sheet: &oxicells_core::ir::Sheet) -> MergeData {
    use std::collections::BTreeMap;
    let mut grid: BTreeMap<u32, BTreeMap<u32, String>> = BTreeMap::new();
    for row in &sheet.rows {
        for cell in &row.cells {
            let text = word_value(cell);
            if !text.is_empty() {
                grid.entry(row.index).or_default().insert(cell.col, text);
            }
        }
    }
    let Some((&head_row, head)) = grid.iter().next() else { return MergeData::default() };
    let first_col = grid.values().filter_map(|r| r.keys().next().copied()).min().unwrap_or(0);
    let last_col = grid.values().filter_map(|r| r.keys().next_back().copied()).max().unwrap_or(0);
    let columns: Vec<String> = (first_col..=last_col)
        .map(|c| head.get(&c).cloned().unwrap_or_else(|| format!("F{}", c - first_col + 1)))
        .collect();
    let rows = grid
        .range(head_row + 1..)
        .map(|(_, r)| (first_col..=last_col).map(|c| r.get(&c).cloned().unwrap_or_default()).collect())
        .collect();
    MergeData { columns, rows }
}

#[cfg(feature = "suite")]
fn word_value(cell: &oxicells_core::ir::Cell) -> String {
    use oxicells_core::ir::CellValue;
    match &cell.value {
        CellValue::Empty => String::new(),
        CellValue::String(s) => s.clone(),
        CellValue::Utf16(u) => String::from_utf16_lossy(u),
        CellValue::Boolean(b) => if *b { "True" } else { "False" }.to_string(),
        CellValue::Error(e) => e.clone(),
        CellValue::Number(n) => {
            let date = cell.style.number_format.as_deref().is_some_and(oxicells_calc::numfmt::looks_like_a_date);
            if date { serial_text(*n) } else { plain_number(*n) }
        }
    }
}

/// A number as Word gets it from Excel: up to 15 significant digits and 11
/// decimals (`=1/3` arrives as `0.33333333333`), without trailing zeros.
#[cfg(feature = "suite")]
fn plain_number(n: f64) -> String {
    if n == n.trunc() && n.abs() < 1e15 {
        return format!("{}", n as i64);
    }
    let magnitude = if n == 0.0 { 0 } else { n.abs().log10().floor() as i32 + 1 };
    let decimals = (15 - magnitude).clamp(0, 11) as usize;
    let s = format!("{:.*}", decimals, n);
    let s = if s.contains('.') { s.trim_end_matches('0').trim_end_matches('.').to_string() } else { s };
    if s == "-0" { "0".to_string() } else { s }
}

/// An Excel date serial as Word shows a date cell: `m/d/yyyy`, a time as
/// `h:mm:ss AM/PM`, both when the serial has both.
#[cfg(feature = "suite")]
fn serial_text(serial: f64) -> String {
    let days = serial.floor() as i64;
    let secs = ((serial - serial.floor()) * 86400.0).round() as i64;
    let date = if days == 60 {
        Some("2/29/1900".to_string())
    } else if days >= 1 {
        // Excel counts 1900-02-29, which never was: from serial 61 on, day 0 is 1899-12-30.
        let base = if days >= 61 { days - 1 } else { days };
        let (y, m, d) = civil_from_days(base - 1 + days_from_civil(1900, 1, 1));
        Some(format!("{m}/{d}/{y}"))
    } else {
        None
    };
    let time = if secs > 0 && secs < 86400 {
        let (h, mi, s) = (secs / 3600, secs / 60 % 60, secs % 60);
        let h12 = if h % 12 == 0 { 12 } else { h % 12 };
        Some(format!("{h12}:{mi:02}:{s:02} {}", if h < 12 { "AM" } else { "PM" }))
    } else {
        None
    };
    match (date, time) {
        (Some(d), Some(t)) => format!("{d} {t}"),
        (Some(d), None) => d,
        (None, Some(t)) => t,
        (None, None) => "12:00:00 AM".to_string(),
    }
}

#[cfg(feature = "suite")]
fn days_from_civil(y: i64, m: i64, d: i64) -> i64 {
    let y = if m <= 2 { y - 1 } else { y };
    let era = y.div_euclid(400);
    let yoe = y - era * 400;
    let doy = (153 * (m + if m > 2 { -3 } else { 9 }) + 2) / 5 + d - 1;
    let doe = yoe * 365 + yoe / 4 - yoe / 100 + doy;
    era * 146097 + doe - 719468
}

#[cfg(feature = "suite")]
fn civil_from_days(z: i64) -> (i64, i64, i64) {
    let z = z + 719468;
    let era = z.div_euclid(146097);
    let doe = z - era * 146097;
    let yoe = (doe - doe / 1460 + doe / 36524 - doe / 146096) / 365;
    let y = yoe + era * 400;
    let doy = doe - (365 * yoe + yoe / 4 - yoe / 100);
    let mp = (5 * doy + 2) / 153;
    let d = doy - (153 * mp + 2) / 5 + 1;
    let m = if mp < 10 { mp + 3 } else { mp - 9 };
    (if m <= 2 { y + 1 } else { y }, m, d)
}

#[cfg(all(test, feature = "suite"))]
mod tests {
    use super::*;

    #[test]
    fn excel_values_come_as_word_reads_them() {
        // Word's merge from data3.xlsx (2026-11-26): 1234.5 / 0.25 / 1/5/2026 /
        // 0.33333333333 / 1:07:00 PM / 12:30:00 AM.
        assert_eq!(plain_number(1234.5), "1234.5");
        assert_eq!(plain_number(0.25), "0.25");
        assert_eq!(plain_number(7.0), "7");
        assert_eq!(plain_number(-30.0), "-30");
        assert_eq!(plain_number(1.0 / 3.0), "0.33333333333");
        assert_eq!(serial_text(46027.0), "1/5/2026");
        assert_eq!(serial_text(46387.0), "12/31/2026");
        assert_eq!(serial_text(13.0 / 24.0 + 7.0 / 1440.0), "1:07:00 PM");
        assert_eq!(serial_text(0.5 / 24.0), "12:30:00 AM");
        assert_eq!(serial_text(60.0), "2/29/1900");
        assert_eq!(serial_text(61.0), "3/1/1900");
    }
}
