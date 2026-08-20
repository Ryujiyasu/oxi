// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! Oxi XLSX Renderer — draws a worksheet for pixel comparison against Excel.
//!
//! Excel's truth comes from `ExportAsFixedFormat` (a PDF), rasterised
//! alongside this. What is drawn here is the printed sheet: the used range,
//! cell by cell, with the text each cell shows under its number format.
//!
//! Usage:
//!   oxi-xlsx-renderer <input.xlsx> <output.png> [dpi]
//!
//! Geometry, measured from Excel:
//!   - A column's width is stated in characters; Excel renders the default
//!     8.38 characters as 54 pixels at 96 DPI.
//!   - A row's height is stated in points, 18.75 by default.

use oxicells_core::ir::{CellStyle, CellValue, Sheet, Workbook};
use oxicells_core::parser::parse_xlsx;

/// Excel reports a column's width in characters and renders the default 8.38
/// of them as 54 pixels at 96 DPI, which is 6.4488 pixels per character.
const PIXELS_PER_CHARACTER: f32 = 54.0 / 8.38;
const DEFAULT_ROW_POINTS: f32 = 18.75;

struct Geometry {
    /// Left edge of each column, and the sheet's full width.
    columns: Vec<f32>,
    /// Top edge of each row, and the sheet's full height.
    rows: Vec<f32>,
    first_column: u32,
    first_row: u32,
}

fn used_extent(sheet: &Sheet) -> (u32, u32, u32, u32) {
    let mut first_row = u32::MAX;
    let mut last_row = 0;
    let mut first_column = u32::MAX;
    let mut last_column = 0;
    for row in &sheet.rows {
        for cell in &row.cells {
            if matches!(cell.value, CellValue::Empty) && cell.formula.is_none() {
                continue;
            }
            first_row = first_row.min(row.index);
            last_row = last_row.max(row.index);
            first_column = first_column.min(cell.col);
            last_column = last_column.max(cell.col);
        }
    }
    if first_row == u32::MAX {
        return (1, 1, 0, 0);
    }
    (first_row, first_column, last_row, last_column)
}

fn geometry(sheet: &Sheet, scale: f32) -> Geometry {
    let (first_row, first_column, last_row, last_column) = used_extent(sheet);

    let mut columns = vec![0.0];
    for column in first_column..=last_column {
        let characters = sheet
            .col_widths
            .get(column as usize)
            .copied()
            .filter(|width| *width > 0.0)
            .unwrap_or(sheet.default_col_width);
        let hidden = sheet.hidden_cols.contains(&column);
        let width = if hidden {
            0.0
        } else {
            (characters * PIXELS_PER_CHARACTER * scale).round()
        };
        columns.push(columns.last().unwrap() + width);
    }

    let mut rows = vec![0.0];
    for index in first_row..=last_row {
        let held = sheet.rows.iter().find(|row| row.index == index);
        let hidden = held.is_some_and(|row| row.hidden);
        let points = held
            .and_then(|row| row.height)
            .unwrap_or(if sheet.default_row_height > 0.0 {
                sheet.default_row_height
            } else {
                DEFAULT_ROW_POINTS
            });
        let height = if hidden {
            0.0
        } else {
            (points * scale * 96.0 / 72.0).round()
        };
        rows.push(rows.last().unwrap() + height);
    }

    Geometry {
        columns,
        rows,
        first_column,
        first_row,
    }
}

fn main() {
    let args: Vec<String> = std::env::args().collect();
    if args.len() < 3 {
        eprintln!("Usage: {} <input.xlsx> <output.png> [dpi]", args[0]);
        std::process::exit(1);
    }
    let dpi: f32 = args.get(3).and_then(|value| value.parse().ok()).unwrap_or(96.0);
    let scale = dpi / 96.0;

    let data = match std::fs::read(&args[1]) {
        Ok(data) => data,
        Err(error) => {
            eprintln!("cannot read {}: {error}", args[1]);
            std::process::exit(1);
        }
    };
    let workbook: Workbook = match parse_xlsx(&data) {
        Ok(workbook) => workbook,
        Err(error) => {
            eprintln!("cannot parse {}: {error}", args[1]);
            std::process::exit(1);
        }
    };
    let Some(sheet) = workbook.sheets.first() else {
        eprintln!("the workbook holds no sheets");
        std::process::exit(1);
    };

    let layout = geometry(sheet, scale);
    let width = *layout.columns.last().unwrap_or(&0.0) as u32;
    let height = *layout.rows.last().unwrap_or(&0.0) as u32;
    if width == 0 || height == 0 {
        eprintln!("the sheet has nothing in it to draw");
        std::process::exit(1);
    }

    let canvas = draw(sheet, &layout, width, height, scale);
    if let Err(error) = canvas.save(&args[2]) {
        eprintln!("cannot write {}: {error}", args[2]);
        std::process::exit(1);
    }
    println!("{} {}x{} @{}dpi", args[2], width, height, dpi);
}

#[cfg(windows)]
fn draw(
    sheet: &Sheet,
    layout: &Geometry,
    width: u32,
    height: u32,
    scale: f32,
) -> image::RgbImage {
    windows_draw::draw(sheet, layout, width, height, scale)
}

#[cfg(not(windows))]
fn draw(
    _sheet: &Sheet,
    _layout: &Geometry,
    width: u32,
    height: u32,
    _scale: f32,
) -> image::RgbImage {
    // Text needs the platform's own rasteriser to match Excel, so elsewhere
    // this only reports the geometry it would have drawn into.
    image::RgbImage::from_pixel(width, height, image::Rgb([255, 255, 255]))
}

/// The text a cell shows, under whatever number format it wears.
fn cell_text(value: &CellValue, style: &CellStyle) -> String {
    match value {
        CellValue::Empty => String::new(),
        CellValue::String(text) => text.clone(),
        CellValue::Boolean(true) => "TRUE".to_string(),
        CellValue::Boolean(false) => "FALSE".to_string(),
        CellValue::Error(text) => text.clone(),
        CellValue::Number(number) => oxicells_core::format_number(
            *number,
            style.number_format.as_deref().unwrap_or("General"),
        ),
    }
}

/// Where the text sits across a cell. Excel puts numbers to the right and text
/// to the left unless the cell says otherwise.
fn alignment(style: &CellStyle, value: &CellValue) -> Align {
    match style.horizontal_align.as_deref() {
        Some("center") => Align::Centre,
        Some("right") => Align::Right,
        Some("left") => Align::Left,
        _ => match value {
            CellValue::Number(_) => Align::Right,
            CellValue::Boolean(_) => Align::Centre,
            _ => Align::Left,
        },
    }
}

#[derive(Clone, Copy)]
enum Align {
    Left,
    Centre,
    Right,
}

#[cfg(windows)]
mod windows_draw {
    use super::{alignment, cell_text, Align, Geometry};
    use oxicells_core::ir::{CellValue, Sheet};
    use windows::core::PCWSTR;
    use windows::Win32::Foundation::{COLORREF, RECT};
    use windows::Win32::Graphics::Gdi::*;

    fn wide(text: &str) -> Vec<u16> {
        text.encode_utf16().chain(std::iter::once(0)).collect()
    }

    fn colour(hex: Option<&str>, fallback: u32) -> COLORREF {
        let value = hex
            .and_then(|hex| u32::from_str_radix(hex.trim_start_matches('#'), 16).ok())
            .unwrap_or(fallback);
        // A sheet writes RRGGBB; GDI wants BGR.
        let (r, g, b) = ((value >> 16) & 0xFF, (value >> 8) & 0xFF, value & 0xFF);
        COLORREF(b << 16 | g << 8 | r)
    }

    pub fn draw(
        sheet: &Sheet,
        layout: &Geometry,
        width: u32,
        height: u32,
        scale: f32,
    ) -> image::RgbImage {
        unsafe {
            let screen = GetDC(None);
            let dc = CreateCompatibleDC(screen);
            let mut bits: *mut std::ffi::c_void = std::ptr::null_mut();
            let info = BITMAPINFO {
                bmiHeader: BITMAPINFOHEADER {
                    biSize: std::mem::size_of::<BITMAPINFOHEADER>() as u32,
                    biWidth: width as i32,
                    biHeight: -(height as i32),
                    biPlanes: 1,
                    biBitCount: 32,
                    biCompression: BI_RGB.0,
                    ..Default::default()
                },
                ..Default::default()
            };
            let bitmap =
                CreateDIBSection(dc, &info, DIB_RGB_COLORS, &mut bits, None, 0).unwrap();
            let previous = SelectObject(dc, bitmap);

            let whole = RECT {
                left: 0,
                top: 0,
                right: width as i32,
                bottom: height as i32,
            };
            let white = CreateSolidBrush(COLORREF(0x00FF_FFFF));
            FillRect(dc, &whole, white);
            let _ = DeleteObject(white);

            SetBkMode(dc, TRANSPARENT);

            for row in &sheet.rows {
                if row.index < layout.first_row {
                    continue;
                }
                let top_at = (row.index - layout.first_row) as usize;
                let (Some(top), Some(bottom)) =
                    (layout.rows.get(top_at), layout.rows.get(top_at + 1))
                else {
                    continue;
                };
                if bottom <= top {
                    continue; // a hidden row takes no space
                }
                for cell in &row.cells {
                    if cell.col < layout.first_column {
                        continue;
                    }
                    let left_at = (cell.col - layout.first_column) as usize;
                    let (Some(left), Some(right)) =
                        (layout.columns.get(left_at), layout.columns.get(left_at + 1))
                    else {
                        continue;
                    };
                    if right <= left {
                        continue;
                    }
                    let box_ = RECT {
                        left: *left as i32,
                        top: *top as i32,
                        right: *right as i32,
                        bottom: *bottom as i32,
                    };

                    if let Some(fill) = cell.style.bg_color.as_deref() {
                        let brush = CreateSolidBrush(colour(Some(fill), 0xFFFFFF));
                        FillRect(dc, &box_, brush);
                        let _ = DeleteObject(brush);
                    }

                    let text = cell_text(&cell.value, &cell.style);
                    if text.is_empty() {
                        continue;
                    }
                    let points = cell.style.font_size.unwrap_or(11.0);
                    let pixels = -((points * scale * 96.0 / 72.0).round() as i32);
                    // A cell names its own typeface; Calibri is only the
                    // fallback for one that does not.
                    let face = wide(
                        cell.style.font_name.as_deref().unwrap_or("Calibri"),
                    );
                    let font = CreateFontW(
                        pixels,
                        0,
                        0,
                        0,
                        if cell.style.bold { 700 } else { 400 },
                        u32::from(cell.style.italic),
                        0,
                        0,
                        DEFAULT_CHARSET.0 as u32,
                        OUT_DEFAULT_PRECIS.0 as u32,
                        CLIP_DEFAULT_PRECIS.0 as u32,
                        // Excel prints greyscale-antialiased glyphs. ClearType
                        // would colour their edges, which reads as a horizontal
                        // shift once the comparison turns both to grey.
                        ANTIALIASED_QUALITY.0 as u32,
                        (DEFAULT_PITCH.0 | FF_DONTCARE.0) as u32,
                        PCWSTR(face.as_ptr()),
                    );
                    let previous_font = SelectObject(dc, font);
                    SetTextColor(dc, colour(cell.style.font_color.as_deref(), 0x000000));

                    // Excel keeps a small gutter either side of a cell's text.
                    let gutter = (2.0 * scale).round() as i32;
                    let mut area = box_;
                    area.left += gutter;
                    area.right -= gutter;
                    let format = match alignment(&cell.style, &cell.value) {
                        Align::Left => DT_LEFT,
                        Align::Centre => DT_CENTER,
                        Align::Right => DT_RIGHT,
                    } | DT_VCENTER
                        | DT_SINGLELINE
                        | DT_NOPREFIX;
                    let mut body = wide(&text);
                    body.pop();
                    DrawTextW(dc, &mut body, &mut area, format);

                    SelectObject(dc, previous_font);
                    let _ = DeleteObject(font);
                }
            }

            // Excel prints no gridlines by default, so none are drawn here.
            let mut canvas = image::RgbImage::new(width, height);
            let pixels =
                std::slice::from_raw_parts(bits as *const u8, (width * height * 4) as usize);
            for y in 0..height {
                for x in 0..width {
                    let at = ((y * width + x) * 4) as usize;
                    canvas.put_pixel(
                        x,
                        y,
                        image::Rgb([pixels[at + 2], pixels[at + 1], pixels[at]]),
                    );
                }
            }

            SelectObject(dc, previous);
            let _ = DeleteObject(bitmap);
            let _ = DeleteDC(dc);
            ReleaseDC(None, screen);
            let _ = CellValue::Empty;
            canvas
        }
    }
}
