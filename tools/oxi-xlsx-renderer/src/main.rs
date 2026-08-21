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

use oxicells_core::ir::{CellStyle, CellValue, Row, Sheet, Workbook};
use oxicells_core::parser::parse_xlsx;

/// A stored column width already carries the gutter either side of a cell's
/// text, so it is not the width a person types into Excel: typing 10 stores
/// 10.625. Pixels come back out of it by OOXML's own rule, which a ruled
/// worksheet confirmed to the pixel.
fn column_pixels(width: f32, digit: f32) -> f32 {
    let padding = (128.0 / digit).trunc();
    (((256.0 * width + padding) / 256.0) * digit).trunc()
}
/// What a digit measures when the standard font could not be asked: Calibri 11,
/// which is what a workbook that names no font gets.
const FALLBACK_DIGIT_WIDTH: f32 = 7.0;
const DEFAULT_ROW_POINTS: f32 = 18.75;
/// Excel's own default column width for a sheet that states none, as a plain
/// count of digits.
const DEFAULT_CHARACTERS: f32 = 8.43;
/// The gutter Excel keeps either side of a cell's text, in pixels. A width the
/// sheet states already has it folded in; a default width does not.
const COLUMN_GUTTER: f32 = 5.0;

pub(crate) struct Geometry {
    /// Left edge of each column, and the sheet's full width.
    columns: Vec<f32>,
    /// Top edge of each row, and the sheet's full height.
    rows: Vec<f32>,
    first_column: u32,
    first_row: u32,
}

fn used_extent(sheet: &Sheet, plain: &CellStyle) -> (u32, u32, u32, u32) {
    let mut first_row = u32::MAX;
    let mut last_row = 0;
    let mut first_column = u32::MAX;
    let mut last_column = 0;
    for row in &sheet.rows {
        for cell in &row.cells {
            // A cell counts as used when it holds something OR when it is
            // dressed differently from the workbook's default — a ruled but
            // empty cell is part of the range Excel hands over, and leaving it
            // out shifts the whole sheet against Excel's own picture.
            let empty = matches!(cell.value, CellValue::Empty) && cell.formula.is_none();
            if empty && &cell.style == plain {
                continue;
            }
            first_row = first_row.min(row.index);
            last_row = last_row.max(row.index);
            first_column = first_column.min(cell.col);
            last_column = last_column.max(cell.col);
        }
    }
    // A sheet says how far it reaches, and that can be further than its last
    // filled cell. Excel hands over the declared range when asked for a
    // picture, so text running on past the last cell has somewhere to go.
    if let Some((start_row, start_column, end_row, end_column)) = sheet.declared_range {
        if first_row == u32::MAX {
            return (start_row, start_column, end_row, end_column);
        }
        return (
            first_row.min(start_row),
            first_column.min(start_column),
            last_row.max(end_row),
            last_column.max(end_column),
        );
    }
    if first_row == u32::MAX {
        return (1, 1, 0, 0);
    }
    (first_row, first_column, last_row, last_column)
}

fn geometry(sheet: &Sheet, scale: f32, digit_width: f32, plain: &CellStyle) -> Geometry {
    let (first_row, first_column, last_row, last_column) = used_extent(sheet, plain);

    let mut columns = vec![0.0];
    for column in first_column..=last_column {
        let stated = sheet
            .col_widths
            .get(column as usize)
            .copied()
            .filter(|width| *width > 0.0);
        let hidden = sheet.hidden_cols.contains(&column);
        let width = if hidden {
            0.0
        } else {
            match stated {
                // A stated width already carries the gutter either side of the
                // cell's text.
                Some(width) => column_pixels(width, digit_width) * scale,
                // A default the sheet states is measured the same way. Excel's
                // own default, for a sheet that states none, is a plain count
                // of characters, so there the gutter is added here instead.
                None if sheet.default_col_width > 0.0 => {
                    column_pixels(sheet.default_col_width, digit_width) * scale
                }
                None => (DEFAULT_CHARACTERS * digit_width + COLUMN_GUTTER).trunc() * scale,
            }
        };
        columns.push(columns.last().unwrap() + width);
    }

    let mut rows = vec![0.0];
    for index in first_row..=last_row {
        let held = sheet.rows.iter().find(|row| row.index == index);
        let hidden = held.is_some_and(|row| row.hidden);
        let points = held.and_then(|row| row.height).unwrap_or_else(|| {
            // A row Excel draws is a whole number of pixels tall, and a pixel
            // is 0.75pt, so the height a sheet states as its default is rounded
            // UP to the next 0.75 before it is used: a sheet saying 13 is drawn
            // at 13.5, which is 18px. A sheet already stating a multiple — 18.75
            // for the usual 11pt font — is left where it is.
            let stated = if sheet.default_row_height > 0.0 {
                sheet.default_row_height
            } else {
                DEFAULT_ROW_POINTS
            };
            (stated / 0.75).ceil() * 0.75
        });
        let height = if hidden {
            0.0
        } else {
            // A height in points becomes whole pixels by truncation: 20.1pt
            // is 26.8px and Excel draws 26.
            (points * scale * 96.0 / 72.0).trunc()
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

    // A column is measured in digits of the workbook's standard font, so the
    // sheet cannot be laid out until that digit has been measured.
    let digit_width = digit_width(&workbook.default_style).unwrap_or(FALLBACK_DIGIT_WIDTH);
    let layout = geometry(sheet, scale, digit_width, &workbook.default_style);
    // One pixel past the last edge, because a rule on that edge is drawn there
    // and Excel's own picture is that pixel wider and taller.
    let width = *layout.columns.last().unwrap_or(&0.0) as u32 + 1;
    let height = *layout.rows.last().unwrap_or(&0.0) as u32 + 1;
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

/// How wide a digit is in the font a workbook treats as standard.
#[cfg(windows)]
fn digit_width(style: &CellStyle) -> Option<f32> {
    windows_draw::digit_width(style)
}

#[cfg(not(windows))]
fn digit_width(_style: &CellStyle) -> Option<f32> {
    None
}

#[cfg(windows)]
fn draw(
    sheet: &Sheet,
    layout: &Geometry,
    width: u32,
    height: u32,
    scale: f32,
) -> image::RgbImage {
    // Excel draws its own text through DirectWrite, so that is the path that
    // can match it. GDI stays reachable while the two are being compared.
    if std::env::var("OXI_XLSX_GDI").is_err() {
        match dwrite_draw::draw(sheet, layout, width, height, scale) {
            Ok(canvas) => return canvas,
            Err(error) => eprintln!("DirectWrite could not draw the sheet: {error}"),
        }
    }
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
pub(crate) fn cell_text(value: &CellValue, style: &CellStyle) -> String {
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
/// How a table dresses one of its cells: Excel paints this, and no cell inside
/// the range carries any of it in its own style.
pub(crate) struct Dressed {
    fill: Option<String>,
    /// A header row is written in white on the accent colour.
    white_text: bool,
}

/// A table's header carries a filter button in every column. Measured off a
/// worksheet Excel drew: 17 by 17 pixels against the right edge of the cell,
/// eight below its top, a pale face inside a grey outline, and a seven-wide
/// triangle narrowing to a point four rows down.
pub(crate) const FILTER_BUTTON: i32 = 17;
pub(crate) const FILTER_BUTTON_TOP: i32 = 8;

/// Whether the cell at this spot carries a filter button.
pub(crate) fn has_filter_button(sheet: &Sheet, row: u32, column: u32) -> bool {
    sheet.tables.iter().any(|table| {
        table.header_rows > 0
            && row >= table.start_row
            && row < table.start_row + table.header_rows
            && column >= table.start_col
            && column <= table.end_col
    })
}

pub(crate) fn dressed_by_table(sheet: &Sheet, row: u32, column: u32) -> Option<Dressed> {
    let table = sheet.tables.iter().find(|table| {
        row >= table.start_row
            && row <= table.end_row
            && column >= table.start_col
            && column <= table.end_col
    })?;
    if row < table.start_row + table.header_rows {
        return Some(Dressed {
            fill: table.accent.clone(),
            white_text: true,
        });
    }
    if !table.banded_rows {
        return None;
    }
    // The first row below the header is banded, then every other one.
    let below = row - (table.start_row + table.header_rows);
    Some(Dressed {
        fill: if below % 2 == 0 { table.band.clone() } else { None },
        white_text: false,
    })
}

/// What a merge does to a cell: the top-left one is drawn across the whole
/// block, and the rest are not drawn at all.
pub(crate) enum Merged {
    /// Drawn across this many further columns and rows.
    Anchor { columns: u32, rows: u32 },
    Covered,
}

pub(crate) fn merges(sheet: &Sheet) -> std::collections::HashMap<(u32, u32), Merged> {
    let mut held = std::collections::HashMap::new();
    for merge in &sheet.merge_cells {
        for row in merge.start_row..=merge.end_row {
            for column in merge.start_col..=merge.end_col {
                let entry = if row == merge.start_row && column == merge.start_col {
                    Merged::Anchor {
                        columns: merge.end_col - merge.start_col,
                        rows: merge.end_row - merge.start_row,
                    }
                } else {
                    Merged::Covered
                };
                held.insert((row, column), entry);
            }
        }
    }
    held
}

/// Whether a cell holds nothing, so a neighbour's text may run over it.
fn is_free(
    row: &Row,
    column: u32,
    merged: &std::collections::HashMap<(u32, u32), Merged>,
) -> bool {
    if merged.contains_key(&(row.index, column)) {
        return false;
    }
    !row.cells.iter().any(|cell| {
        cell.col == column && !cell_text(&cell.value, &cell.style).is_empty()
    })
}

/// How far left of its own cell a run-on may reach, in pixels, stopping at the
/// first neighbour that holds something.
pub(crate) fn room_before(
    layout: &Geometry,
    row: &Row,
    column: u32,
    wanted: i32,
    merged: &std::collections::HashMap<(u32, u32), Merged>,
) -> i32 {
    let mut room = 0;
    let mut at = column;
    while room < wanted && at > layout.first_column {
        at -= 1;
        if !is_free(row, at, merged) {
            break;
        }
        let index = (at - layout.first_column) as usize;
        let (Some(left), Some(right)) =
            (layout.columns.get(index), layout.columns.get(index + 1))
        else {
            break;
        };
        room += (right - left) as i32;
    }
    room.min(wanted)
}

/// The same, rightward.
pub(crate) fn room_after(
    layout: &Geometry,
    row: &Row,
    column: u32,
    wanted: i32,
    merged: &std::collections::HashMap<(u32, u32), Merged>,
) -> i32 {
    let mut room = 0;
    let mut at = column;
    loop {
        if room >= wanted {
            break;
        }
        at += 1;
        let index = (at - layout.first_column) as usize;
        let (Some(left), Some(right)) =
            (layout.columns.get(index), layout.columns.get(index + 1))
        else {
            break;
        };
        if !is_free(row, at, merged) {
            break;
        }
        room += (right - left) as i32;
    }
    room.min(wanted)
}

/// How Excel draws one kind of rule, measured off a worksheet holding one of
/// each: how far the ink reaches either side of the boundary, and which pixels
/// along it are inked.
pub(crate) struct Rule {
    /// Pixels before the boundary the ink starts at, and after it ends.
    before: i32,
    after: i32,
    /// The pixel at the boundary itself is skipped by a double rule.
    hollow: bool,
    /// A run length and a period: `Some((3, 4))` inks three of every four.
    dashes: Option<(i32, i32)>,
}

pub(crate) fn rule_for(kind: &str) -> Rule {
    let solid = |before, after| Rule { before, after, hollow: false, dashes: None };
    match kind {
        "medium" | "mediumDashed" | "mediumDashDot" | "mediumDashDotDot" => solid(1, 0),
        "thick" => solid(1, 1),
        "double" => Rule { before: 1, after: 1, hollow: true, dashes: None },
        "hair" => Rule { before: 0, after: 0, hollow: false, dashes: Some((1, 2)) },
        "dotted" => Rule { before: 0, after: 0, hollow: false, dashes: Some((2, 4)) },
        "dashed" | "dashDot" | "dashDotDot" | "slantDashDot" => {
            Rule { before: 0, after: 0, hollow: false, dashes: Some((3, 4)) }
        }
        // "thin" and anything unfamiliar: a single pixel on the boundary.
        _ => solid(0, 0),
    }
}

pub(crate) fn alignment(style: &CellStyle, value: &CellValue) -> Align {
    match style.horizontal_align.as_deref() {
        Some("center") => Align::Centre,
        Some("right") => Align::Right,
        Some("left") => Align::Left,
        // A heading centred across its neighbours is centred within the cell
        // that holds the text.
        Some("centerContinuous") => Align::Centre,
        // Spread out to fill the cell, which is how Japanese workbooks set a
        // title: 第 ３ 表, not 第３表.
        Some("distributed") | Some("justify") => Align::Spread,
        _ => match value {
            CellValue::Number(_) => Align::Right,
            CellValue::Boolean(_) => Align::Centre,
            _ => Align::Left,
        },
    }
}

#[derive(Clone, Copy, PartialEq)]
pub(crate) enum Align {
    Left,
    Centre,
    Right,
    /// Spread across the width of the cell, letters and all.
    Spread,
}

#[cfg(windows)]
mod dwrite_draw;

#[cfg(windows)]
mod windows_draw {
    use super::{alignment, cell_text, Align, Geometry, Merged};
    use oxicells_core::ir::{BorderLine, CellStyle, CellValue, Row, Sheet};
    use windows::core::PCWSTR;
    use windows::Win32::Foundation::{COLORREF, RECT, SIZE};
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

    /// Measures a digit of the workbook's standard font, which is the unit a
    /// column width is stated in.
    pub fn digit_width(style: &CellStyle) -> Option<f32> {
        unsafe {
            let dc = CreateCompatibleDC(None);
            if dc.is_invalid() {
                return None;
            }
            let points = style.font_size.unwrap_or(11.0);
            let face = wide(style.font_name.as_deref().unwrap_or("Calibri"));
            let font = CreateFontW(
                -((points * 96.0 / 72.0).round() as i32),
                0,
                0,
                0,
                400,
                0,
                0,
                0,
                DEFAULT_CHARSET.0 as u32,
                OUT_DEFAULT_PRECIS.0 as u32,
                CLIP_DEFAULT_PRECIS.0 as u32,
                ANTIALIASED_QUALITY.0 as u32,
                (DEFAULT_PITCH.0 | FF_DONTCARE.0) as u32,
                PCWSTR(face.as_ptr()),
            );
            let previous = SelectObject(dc, font);
            let mut size = SIZE::default();
            let zero = wide("0");
            let measured = GetTextExtentPoint32W(dc, &zero[..1], &mut size).as_bool();
            SelectObject(dc, previous);
            let _ = DeleteObject(font);
            let _ = DeleteDC(dc);
            measured.then_some(size.cx as f32)
        }
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

            let merged = super::merges(sheet);
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
                    // A merged block belongs to its top-left cell; the cells
                    // it covers are not drawn at all.
                    let (spans_columns, spans_rows) =
                        match merged.get(&(row.index, cell.col)) {
                            Some(Merged::Covered) => continue,
                            Some(Merged::Anchor { columns, rows }) => (*columns, *rows),
                            None => (0, 0),
                        };
                    let left_at = (cell.col - layout.first_column) as usize;
                    let (Some(left), Some(right)) = (
                        layout.columns.get(left_at),
                        layout.columns.get(left_at + 1 + spans_columns as usize),
                    ) else {
                        continue;
                    };
                    if right <= left {
                        continue;
                    }
                    let bottom = layout
                        .rows
                        .get(top_at + 1 + spans_rows as usize)
                        .unwrap_or(bottom);
                    let box_ = RECT {
                        left: *left as i32,
                        top: *top as i32,
                        right: *right as i32,
                        bottom: *bottom as i32,
                    };

                    // A table's own dress goes down first; a cell that names a
                    // fill of its own paints over it.
                    let dress = super::dressed_by_table(sheet, row.index, cell.col);
                    let fill = cell
                        .style
                        .bg_color
                        .as_deref()
                        .or_else(|| dress.as_ref().and_then(|d| d.fill.as_deref()));
                    if let Some(fill) = fill {
                        let brush = CreateSolidBrush(colour(Some(fill), 0xFFFFFF));
                        FillRect(dc, &box_, brush);
                        let _ = DeleteObject(brush);
                    }

                    // A rule sits ON the boundary, not inside the cell: the
                    // bottom of one row and the top of the next are the same
                    // pixel. Measured against Excel on a box ruled all the way
                    // round, whose edges landed exactly on the column and row
                    // starts.
                    let edges: [(&Option<BorderLine>, bool, i32); 4] = [
                        (&cell.style.border_top, true, box_.top),
                        (&cell.style.border_bottom, true, box_.bottom),
                        (&cell.style.border_left, false, box_.left),
                        (&cell.style.border_right, false, box_.right),
                    ];
                    for (line, horizontal, at) in edges {
                        let Some(line) = line else { continue };
                        let rule = super::rule_for(&line.style);
                        let ink = CreateSolidBrush(colour(line.color.as_deref(), 0x000000));
                        for step in -rule.before..=rule.after {
                            if rule.hollow && step == 0 {
                                continue;
                            }
                            let edge = if horizontal {
                                RECT { top: at + step, bottom: at + step + 1, ..box_ }
                            } else {
                                RECT { left: at + step, right: at + step + 1, ..box_ }
                            };
                            match rule.dashes {
                                None => {
                                    FillRect(dc, &edge, ink);
                                }
                                Some((on, period)) => {
                                    // A broken rule is inked run by run, so the
                                    // gaps fall where Excel puts them.
                                    let (start, stop) = if horizontal {
                                        (edge.left, edge.right)
                                    } else {
                                        (edge.top, edge.bottom)
                                    };
                                    let mut at_run = start;
                                    while at_run < stop {
                                        let run_end = (at_run + on).min(stop);
                                        let piece = if horizontal {
                                            RECT { left: at_run, right: run_end, ..edge }
                                        } else {
                                            RECT { top: at_run, bottom: run_end, ..edge }
                                        };
                                        FillRect(dc, &piece, ink);
                                        at_run += period;
                                    }
                                }
                            }
                        }
                        let _ = DeleteObject(ink);
                    }

                    // A carriage return would otherwise be drawn as a glyph.
                    let filtered = super::has_filter_button(sheet, row.index, cell.col);
                    if filtered {
                        let left = box_.right - super::FILTER_BUTTON;
                        let top = box_.top + super::FILTER_BUTTON_TOP;
                        let face = RECT {
                            left,
                            top,
                            right: box_.right,
                            bottom: top + super::FILTER_BUTTON,
                        };
                        let outline = CreateSolidBrush(colour(Some("A6ACB3"), 0xA6ACB3));
                        FillRect(dc, &face, outline);
                        let _ = DeleteObject(outline);
                        let pale = CreateSolidBrush(colour(Some("FEFEFE"), 0xFEFEFE));
                        let inside = RECT {
                            left: face.left + 1,
                            top: face.top + 1,
                            right: face.right - 1,
                            bottom: face.bottom - 1,
                        };
                        FillRect(dc, &inside, pale);
                        let _ = DeleteObject(pale);
                        // Seven pixels wide at the top, losing one either side
                        // each row until it is a single point.
                        let arrow = CreateSolidBrush(colour(Some("58595B"), 0x58595B));
                        let middle = (face.left + face.right) / 2;
                        for step in 0..4 {
                            let half = 3 - step;
                            let bar = RECT {
                                left: middle - half,
                                top: face.top + 7 + step,
                                right: middle + half + 1,
                                bottom: face.top + 8 + step,
                            };
                            FillRect(dc, &bar, arrow);
                        }
                        let _ = DeleteObject(arrow);
                    }

                    let text = cell_text(&cell.value, &cell.style)
                        .replace("\r\n", "\n");
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
                        u32::from(cell.style.underline),
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
                    // A table header is written in white on the accent, and
                    // that beats whatever the cell's own format inherited.
                    let header = matches!(&dress, Some(dress) if dress.white_text);
                    SetTextColor(
                        dc,
                        if header {
                            colour(None, 0x00FF_FFFF)
                        } else {
                            colour(cell.style.font_color.as_deref(), 0x0000_0000)
                        },
                    );

                    // Excel keeps a small gutter either side of a cell's text.
                    let gutter = (2.0 * scale).round() as i32;
                    let mut area = box_;
                    area.left += gutter;
                    area.right -= gutter;
                    if filtered {
                        area.right -= super::FILTER_BUTTON;
                    }
                    let placed = alignment(&cell.style, &cell.value);
                    let mut body = wide(&text);
                    body.pop();

                    // Text too long for its cell runs on over the neighbours,
                    // as long as they are empty — that is what Excel shows, and
                    // a wrapping cell keeps to itself instead.
                    // Only text runs on. A number that will not fit stays in
                    // its cell — Excel shows ##### rather than let it spill.
                    let runs_on = !cell.style.wrap_text
                        && placed != Align::Spread
                        && matches!(cell.value, CellValue::String(_));
                    if runs_on {
                        let mut measured = SIZE::default();
                        if GetTextExtentPoint32W(dc, &body, &mut measured).as_bool()
                            && measured.cx > area.right - area.left
                        {
                            let spare = measured.cx - (area.right - area.left);
                            let (leftward, rightward) = match placed {
                                Align::Left | Align::Spread => (0, spare),
                                Align::Right => (spare, 0),
                                Align::Centre => (spare / 2, spare - spare / 2),
                            };
                            area.left -=
                                super::room_before(layout, row, cell.col, leftward, &merged);
                            // A merged block's own columns are already inside
                            // the box, so the search for room starts past them.
                            // A merged block's own columns are already inside
                            // the box, so the search for room starts past them.
                            let after = super::room_after(
                                layout,
                                row,
                                cell.col + spans_columns,
                                rightward,
                                &merged,
                            );
                            area.right += after;
                        }
                    }

                    let placed_flag = match placed {
                        Align::Left => DT_LEFT,
                        // GDI has no spread; centring is the closest it gets.
                        Align::Centre | Align::Spread => DT_CENTER,
                        Align::Right => DT_RIGHT,
                    } | DT_NOPREFIX;
                    // A cell can hold its own breaks, put there with alt+enter,
                    // whatever it says about wrapping.
                    if cell.style.wrap_text || text.contains('\n') {
                        let format = if cell.style.wrap_text {
                            placed_flag | DT_WORDBREAK
                        } else {
                            placed_flag
                        };
                        // Several lines need their own height measured before
                        // they can be placed; DT_VCENTER only centres one.
                        let mut measured = area;
                        let mut probe = body.clone();
                        DrawTextW(dc, &mut probe, &mut measured, format | DT_CALCRECT);
                        let slack =
                            (area.bottom - area.top) - (measured.bottom - measured.top);
                        if slack > 0 {
                            // Where the block of lines sits is the cell's own
                            // rule; Excel leaves a cell at the bottom when it
                            // says nothing.
                            area.top += match cell.style.vertical_align.as_deref() {
                                Some("top") => 0,
                                Some("center") | Some("centre") => slack / 2,
                                _ => slack,
                            };
                        }
                        DrawTextW(dc, &mut body, &mut area, format);
                    } else {
                        // One line sits where the cell says; Excel leaves it
                        // at the bottom when the cell says nothing.
                        let upright = match cell.style.vertical_align.as_deref() {
                            Some("top") => DT_TOP,
                            Some("center") | Some("centre") => DT_VCENTER,
                            _ => DT_BOTTOM,
                        };
                        DrawTextW(
                            dc,
                            &mut body,
                            &mut area,
                            placed_flag | upright | DT_SINGLELINE,
                        );
                    }

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

#[cfg(test)]
mod tests {
    use super::column_pixels;

    /// Every pair here was read off a worksheet Excel drew: the stored width on
    /// the left, the pixels its column occupied on the right, with a digit of
    /// the standard font measuring 8.
    #[test]
    fn a_stored_width_becomes_the_pixels_excel_drew() {
        assert_eq!(column_pixels(10.625, 8.0), 85.0);
        assert_eq!(column_pixels(14.625, 8.0), 117.0);
        assert_eq!(column_pixels(12.625, 8.0), 101.0);
        assert_eq!(column_pixels(9.625, 8.0), 77.0);
    }

    /// The same rule under Calibri, whose digit measures 7: Excel's own default
    /// column comes to 64 pixels.
    #[test]
    fn the_rule_holds_for_a_narrower_digit() {
        assert_eq!(column_pixels(9.140625, 7.0), 64.0);
    }
}
