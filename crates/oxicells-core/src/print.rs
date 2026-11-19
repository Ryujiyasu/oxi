// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! Spreadsheet print pages, laid out from saved values without executing
//! formulas or VBA.
//!
//! The pages are drawn onto a [`PrintCanvas`], which owns the fonts: the
//! command line measures and embeds this machine's font files, the browser
//! build the faces bundled with it. Everything else (print areas, page setup,
//! scaling, breaks, cell drawing) is the same for both.
use crate::ir::{BorderLine, CellStyle, CellValue, MergeCell, Sheet, Visibility, Workbook};
use oxidocs_common::{archive::OoxmlArchive, relationships::parse_relationships};
use oxipdf_core::ir::*;
use quick_xml::{events::Event, Reader};
use std::collections::{BTreeMap, BTreeSet};

type Attrs = BTreeMap<String, String>;
type Area = (u32, u32, u32, u32); // inclusive columns and one-based rows
const MAX_CELLS: u64 = 2_000_000;
const MAX_PAGES: usize = 10_000;

/// Where the print pages are drawn. The canvas picks the font file for each
/// character and measures it, so text is laid out with the widths of the
/// faces that will be embedded.
pub trait PrintCanvas {
    /// Start a new page, `width` × `height` points.
    fn page(&mut self, width: f64, height: f64);
    /// Add a drawing element (paths, clipping, state) to the current page.
    fn element(&mut self, element: ContentElement);
    /// The advance of `text` in points.
    fn width(&self, text: &str, family: &str, size: f64, bold: bool, italic: bool) -> f64;
    /// Draw `text` with its baseline starting at (`x`, `y`).
    #[allow(clippy::too_many_arguments)]
    fn text(&mut self, text: &str, family: &str, size: f64, bold: bool, italic: bool, x: f64, y: f64, color: Color);
}

/// Content of a sheet that the print pages leave out or cannot place
/// faithfully, by kind and the package part it is in.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Unsupported {
    pub kind: String,
    pub part: String,
    pub count: usize,
}

pub fn is_workbook(input: &str, data: &[u8]) -> bool {
    let lower = input.to_ascii_lowercase();
    if [".xlsx", ".xlsm", ".xltx", ".xltm"]
        .iter()
        .any(|ext| lower.ends_with(ext))
    {
        return true;
    }
    // Inspect actual package entries, not bytes inside an embedded workbook.
    OoxmlArchive::new(data).ok().is_some_and(|archive| {
        archive
            .file_names()
            .iter()
            .any(|part| part == "xl/workbook.xml")
    })
}
fn attributes(e: &quick_xml::events::BytesStart<'_>) -> Result<Attrs, String> {
    e.attributes()
        .map(|a| {
            let a = a.map_err(|e| e.to_string())?;
            Ok((
                String::from_utf8_lossy(a.key.as_ref()).into_owned(),
                a.unescape_value().map_err(|e| e.to_string())?.into_owned(),
            ))
        })
        .collect()
}
fn yes(a: &Attrs, key: &str) -> bool {
    a.get(key).is_some_and(|v| v == "1" || v == "true")
}
fn number(a: &Attrs, key: &str, default: f64) -> f64 {
    a.get(key)
        .and_then(|v| v.parse::<f64>().ok())
        .filter(|v| v.is_finite())
        .unwrap_or(default)
}
fn part_path(target: &str) -> String {
    let full = if target.starts_with('/') {
        target.to_owned()
    } else {
        format!("xl/{target}")
    };
    let mut parts = Vec::new();
    for part in full.split('/') {
        match part {
            "" | "." => {}
            ".." => {
                parts.pop();
            }
            _ => parts.push(part),
        }
    }
    parts.join("/")
}
fn sheet_parts(archive: &mut OoxmlArchive) -> Result<BTreeMap<String, String>, String> {
    let rels = parse_relationships(
        &archive
            .read_part("xl/_rels/workbook.xml.rels")
            .map_err(|e| e.to_string())?,
    )
    .map_err(|e| e.to_string())?;
    let xml = archive
        .read_part("xl/workbook.xml")
        .map_err(|e| e.to_string())?;
    let mut reader = Reader::from_str(&xml);
    let mut result = BTreeMap::new();
    loop {
        match reader.read_event().map_err(|e| e.to_string())? {
            Event::Start(e) | Event::Empty(e) if e.local_name().as_ref() == b"sheet" => {
                let a = attributes(&e)?;
                if let (Some(name), Some(rel)) =
                    (a.get("name"), a.get("r:id").and_then(|id| rels.get(id)))
                {
                    if rel.target_mode.as_deref() != Some("External") {
                        result.insert(name.clone(), part_path(&rel.target));
                    }
                }
            }
            Event::Eof => break,
            _ => {}
        }
    }
    Ok(result)
}
#[derive(Clone, Debug)]
struct Settings {
    width: f64,
    height: f64,
    left: f64,
    right: f64,
    top: f64,
    bottom: f64,
    scale: f64,
    fit: bool,
    fit_w: usize,
    fit_h: usize,
    center_x: bool,
    center_y: bool,
    over: bool,
    grid: bool,
    /// pageSetup blackAndWhite: fills left out, text and rules in black.
    mono: bool,
    row_breaks: BTreeSet<u32>,
    col_breaks: BTreeSet<u32>,
    unsupported: BTreeSet<String>,
}
impl Default for Settings {
    fn default() -> Self {
        Self {
            width: 612.,
            height: 792.,
            left: 50.4,
            right: 50.4,
            top: 54.,
            bottom: 54.,
            scale: 1.,
            fit: false,
            fit_w: 1,
            fit_h: 1,
            center_x: false,
            center_y: false,
            over: false,
            grid: false,
            mono: false,
            row_breaks: BTreeSet::new(),
            col_breaks: BTreeSet::new(),
            unsupported: BTreeSet::new(),
        }
    }
}
impl Settings {
    fn read(xml: &str) -> Result<Self, String> {
        let mut s = Self::default();
        let mut reader = Reader::from_str(xml);
        let mut breaks = String::new();
        loop {
            match reader.read_event().map_err(|e| e.to_string())? {
                Event::Start(e) | Event::Empty(e) => {
                    let tag = e.local_name();
                    let a = attributes(&e)?;
                    match tag.as_ref() {
                        b"pageSetUpPr" => s.fit = yes(&a, "fitToPage"),
                        b"pageMargins" => {
                            s.left = number(&a, "left", 0.7) * 72.;
                            s.right = number(&a, "right", 0.7) * 72.;
                            s.top = number(&a, "top", 0.75) * 72.;
                            s.bottom = number(&a, "bottom", 0.75) * 72.;
                        }
                        b"pageSetup" => {
                            (s.width, s.height) = match number(&a, "paperSize", 1.) as u32 {
                                1 => (612., 792.),
                                5 => (612., 1008.),
                                8 => (297. / 25.4 * 72., 420. / 25.4 * 72.),
                                9 => (210. / 25.4 * 72., 297. / 25.4 * 72.),
                                11 => (148. / 25.4 * 72., 210. / 25.4 * 72.),
                                _ => {
                                    s.unsupported.insert("paper-size".into());
                                    (612., 792.)
                                }
                            };
                            if a.get("orientation").is_some_and(|v| v == "landscape") {
                                std::mem::swap(&mut s.width, &mut s.height);
                            }
                            s.scale = number(&a, "scale", 100.).clamp(10., 400.) / 100.;
                            s.fit_w = number(&a, "fitToWidth", 1.).clamp(0., 10000.) as usize;
                            s.fit_h = number(&a, "fitToHeight", 1.).clamp(0., 10000.) as usize;
                            s.over = a.get("pageOrder").is_some_and(|v| v == "overThenDown");
                            s.mono = yes(&a, "blackAndWhite");
                            if yes(&a, "draft") {
                                s.unsupported.insert("print-color-mode".into());
                            }
                            if a.get("cellComments").is_some_and(|v| v != "none") {
                                s.unsupported.insert("printed-comments".into());
                            }
                            if a.get("errors").is_some_and(|v| v != "displayed") {
                                s.unsupported.insert("print-errors".into());
                            }
                            if a.contains_key("paperWidth") || a.contains_key("paperHeight") {
                                s.unsupported.insert("custom-paper-size".into());
                            }
                        }
                        b"printOptions" => {
                            s.center_x = yes(&a, "horizontalCentered");
                            s.center_y = yes(&a, "verticalCentered");
                            s.grid = yes(&a, "gridLines");
                            if yes(&a, "headings") {
                                s.unsupported.insert("print-headings".into());
                            }
                        }
                        b"rowBreaks" => breaks = "row".into(),
                        b"colBreaks" => breaks = "col".into(),
                        b"brk" if yes(&a, "man") => {
                            let id = number(&a, "id", 0.).clamp(0., 1048576.) as u32;
                            if breaks == "row" {
                                s.row_breaks.insert(id);
                            } else if breaks == "col" {
                                s.col_breaks.insert(id);
                            }
                        }
                        _ => {}
                    }
                }
                Event::End(e) if matches!(e.local_name().as_ref(), b"rowBreaks" | b"colBreaks") => {
                    breaks.clear()
                }
                Event::Eof => break,
                _ => {}
            }
        }
        if [s.left, s.right, s.top, s.bottom].iter().any(|v| *v < 0.)
            || s.width - s.left - s.right <= 0.
            || s.height - s.top - s.bottom <= 0.
        {
            return Err("worksheet has invalid print margins".into());
        }
        Ok(s)
    }
}
fn address(text: &str) -> Option<(u32, u32)> {
    let text = text.trim().replace('$', "");
    let at = text.find(|c: char| c.is_ascii_digit())?;
    let (col, row) = text.split_at(at);
    if col.is_empty() || !col.bytes().all(|b| b.is_ascii_alphabetic()) {
        return None;
    }
    let col = col.bytes().try_fold(0u32, |n, b| {
        n.checked_mul(26)?
            .checked_add(u32::from(b.to_ascii_uppercase() - b'A' + 1))
    })?;
    let row = row.parse::<u32>().ok()?;
    (col > 0 && col <= 16384 && row > 0 && row <= 1048576).then_some((col - 1, row))
}
fn areas(text: &str) -> Result<Vec<Area>, String> {
    let mut quoted = false;
    let mut start = 0;
    let mut parts = Vec::new();
    for (i, ch) in text.char_indices() {
        if ch == '\'' {
            quoted = !quoted;
        }
        if ch == ',' && !quoted {
            parts.push(&text[start..i]);
            start = i + 1;
        }
    }
    parts.push(&text[start..]);
    parts
        .into_iter()
        .map(|part| {
            let part = part.trim().trim_start_matches('=');
            let local = part.rsplit_once('!').map_or(part, |(_, r)| r);
            let (first, last) = local.split_once(':').unwrap_or((local, local));
            let (c1, r1) =
                address(first).ok_or_else(|| format!("unsupported print area: {part}"))?;
            let (c2, r2) =
                address(last).ok_or_else(|| format!("unsupported print area: {part}"))?;
            if c2 < c1 || r2 < r1 {
                return Err(format!("invalid print area: {part}"));
            }
            Ok((c1, r1, c2, r2))
        })
        .collect()
}
fn used(sheet: &Sheet) -> Option<Area> {
    let mut range: Option<Area> = None;
    for row in &sheet.rows {
        for cell in &row.cells {
            if cell.formula.is_none()
                && matches!(cell.value, CellValue::Empty)
                && cell.style.bg_color.is_none()
                && cell.style.border_top.is_none()
                && cell.style.border_bottom.is_none()
                && cell.style.border_left.is_none()
                && cell.style.border_right.is_none()
            {
                continue;
            }
            let v = range.get_or_insert((cell.col, row.index, cell.col, row.index));
            v.0 = v.0.min(cell.col);
            v.1 = v.1.min(row.index);
            v.2 = v.2.max(cell.col);
            v.3 = v.3.max(row.index);
        }
    }
    if let Some(v) = range.as_mut() {
        for m in &sheet.merge_cells {
            v.0 = v.0.min(m.start_col);
            v.1 = v.1.min(m.start_row);
            v.2 = v.2.max(m.end_col);
            v.3 = v.3.max(m.end_row);
        }
    }
    range
}
fn axis(sizes: &[f64], capacity: f64, first: u32, breaks: &BTreeSet<u32>) -> Vec<(usize, usize)> {
    let mut pages = Vec::new();
    let mut start = 0;
    let mut width = 0.;
    for (at, size) in sizes.iter().enumerate() {
        if at > start
            && (width + size > capacity + 0.000001 || breaks.contains(&(first + at as u32)))
        {
            pages.push((start, at));
            start = at;
            width = 0.;
        }
        width += size;
    }
    if start < sizes.len() {
        pages.push((start, sizes.len()));
    }
    pages
}
fn positions(sizes: &[f64]) -> Vec<f64> {
    std::iter::once(0.)
        .chain(sizes.iter().scan(0., |sum, v| {
            *sum += v;
            Some(*sum)
        }))
        .collect()
}
fn rgb(value: Option<&str>, default: Color) -> Color {
    let Some(v) = value else {
        return default;
    };
    let v = v.strip_prefix('#').unwrap_or(v);
    let v = if v.len() == 8 { &v[2..] } else { v };
    if v.len() != 6 {
        return default;
    }
    u32::from_str_radix(v, 16)
        .map(|n| {
            Color::Rgb(
                f64::from((n >> 16) & 255) / 255.,
                f64::from((n >> 8) & 255) / 255.,
                f64::from(n & 255) / 255.,
            )
        })
        .unwrap_or(default)
}
fn rectangle(x: f64, y: f64, w: f64, h: f64) -> Vec<PathOp> {
    vec![
        PathOp::MoveTo(x, y),
        PathOp::LineTo(x + w, y),
        PathOp::LineTo(x + w, y + h),
        PathOp::LineTo(x, y + h),
        PathOp::ClosePath,
    ]
}
fn line(canvas: &mut impl PrintCanvas, x: f64, y: f64, x2: f64, y2: f64, width: f64, color: Color) {
    canvas.element(ContentElement::Path(PathData {
        operations: vec![PathOp::MoveTo(x, y), PathOp::LineTo(x2, y2)],
        stroke: Some(StrokeStyle {
            color,
            width,
            line_cap: LineCap::Butt,
            line_join: LineJoin::Miter,
        }),
        fill: None,
    }));
}
/// A cell border as Excel prints its style: the width, a dash pattern drawn
/// as separate segments (in pixels at 96 dpi, as Excel defines them), or two
/// parallel lines for "double".
fn rule(canvas: &mut impl PrintCanvas, from: (f64, f64), to: (f64, f64), style: &str, scale: f64, color: Color) {
    let px = 0.75 * scale;
    let width = match style {
        "medium" | "mediumDashed" | "mediumDashDot" | "mediumDashDotDot" | "slantDashDot" => 1.5,
        "thick" => 2.25,
        "hair" => 0.25,
        _ => 0.75,
    } * scale;
    let pattern: &[f64] = match style {
        "dotted" => &[1., 1.],
        "hair" => &[1., 1.],
        "dashed" => &[3., 1.],
        "mediumDashed" => &[9., 3.],
        "dashDot" | "mediumDashDot" | "slantDashDot" => &[9., 3., 3., 3.],
        "dashDotDot" | "mediumDashDotDot" => &[9., 3., 3., 3., 3., 3.],
        _ => &[],
    };
    let (dx, dy) = (to.0 - from.0, to.1 - from.1);
    let length = (dx * dx + dy * dy).sqrt();
    if length <= 0. {
        return;
    }
    let (ux, uy) = (dx / length, dy / length);
    let stroke = |operations: Vec<PathOp>, width: f64| ContentElement::Path(PathData {
        operations,
        stroke: Some(StrokeStyle { color, width, line_cap: LineCap::Butt, line_join: LineJoin::Miter }),
        fill: None,
    });
    if style == "double" {
        // Two thin rules a pixel and a half to either side of the edge.
        let (nx, ny) = (-uy * 1.5 * px, ux * 1.5 * px);
        let mut operations = Vec::new();
        for side in [-1., 1.] {
            operations.push(PathOp::MoveTo(from.0 + nx * side, from.1 + ny * side));
            operations.push(PathOp::LineTo(to.0 + nx * side, to.1 + ny * side));
        }
        canvas.element(stroke(operations, 0.5 * scale));
        return;
    }
    if pattern.is_empty() {
        canvas.element(stroke(vec![PathOp::MoveTo(from.0, from.1), PathOp::LineTo(to.0, to.1)], width));
        return;
    }
    let mut operations = Vec::new();
    let (mut at, mut i) = (0., 0);
    while at < length {
        let run = pattern[i % pattern.len()] * px;
        if i % 2 == 0 {
            let end = (at + run).min(length);
            operations.push(PathOp::MoveTo(from.0 + ux * at, from.1 + uy * at));
            operations.push(PathOp::LineTo(from.0 + ux * end, from.1 + uy * end));
        }
        at += run;
        i += 1;
    }
    canvas.element(stroke(operations, width));
}
fn text(value: &CellValue, style: &CellStyle) -> String {
    match value {
        CellValue::Empty => String::new(),
        CellValue::String(v) | CellValue::Error(v) => v.clone(),
        CellValue::Utf16(v) => String::from_utf16_lossy(v),
        CellValue::Boolean(v) => if *v { "TRUE" } else { "FALSE" }.into(),
        CellValue::Number(v) => {
            crate::format_number(*v, style.number_format.as_deref().unwrap_or("General"))
        }
    }
}
#[allow(clippy::too_many_arguments)]
fn cell(
    canvas: &mut impl PrintCanvas,
    value: &CellValue,
    style: &CellStyle,
    x: f64,
    y: f64,
    w: f64,
    h: f64,
    scale: f64,
    grid: bool,
    spill: (f64, f64),
    mono: bool,
) {
    // Printed in black and white, Excel leaves fills out and draws text and
    // rules in black.
    let ink = |color: Option<&str>| if mono { Color::Gray(0.) } else { rgb(color, Color::Gray(0.)) };
    if let Some(bg) = style.bg_color.as_deref().filter(|_| !mono) {
        canvas.element(ContentElement::Path(PathData {
            operations: rectangle(x, y, w, h),
            stroke: None,
            fill: Some(rgb(Some(bg), Color::Gray(1.))),
        }));
    }
    let edges = [
        (style.border_top.as_ref(), x, y, x + w, y),
        (style.border_bottom.as_ref(), x, y + h, x + w, y + h),
        (style.border_left.as_ref(), x, y, x, y + h),
        (style.border_right.as_ref(), x + w, y, x + w, y + h),
    ];
    for (border, x, y, x2, y2) in edges {
        if let Some(border) = border {
            rule(canvas, (x, y), (x2, y2), &border.style, scale, ink(border.color.as_deref()));
        } else if grid {
            line(canvas, x, y, x2, y2, 0.25 * scale, Color::Gray(0.7));
        }
    }
    let text = text(value, style);
    if text.is_empty() || w <= 0. || h <= 0. {
        return;
    }
    let family = style.font_name.as_deref().unwrap_or("Calibri");
    let mut size = f64::from(style.font_size.unwrap_or(11.)).max(1.) * scale;
    let room = (w - 4. * scale).max(0.);
    if style.shrink_to_fit && !style.wrap_text {
        size *= (room
            / canvas
                .width(&text, family, size, style.bold, style.italic)
                .max(0.1))
        .min(1.);
    }
    let mut lines = Vec::new();
    for paragraph in text.split('\n') {
        let mut held = String::new();
        for ch in paragraph.chars() {
            if style.wrap_text
                && !held.is_empty()
                && canvas.width(
                    &format!("{held}{ch}"),
                    family,
                    size,
                    style.bold,
                    style.italic,
                ) > room
            {
                lines.push(held);
                held = String::new();
            }
            held.push(ch);
        }
        lines.push(held);
    }
    let line_h = size * 1.2;
    let height = lines.len() as f64 * line_h;
    let top = match style.vertical_align.as_deref() {
        Some("top") => y + scale,
        Some("center") => y + (h - height) / 2.,
        _ => y + h - height - scale,
    };
    // Text runs on over the empty cells beside it (`spill`, left and right),
    // as Excel prints it; it is cut where they end.
    canvas.element(ContentElement::SaveState);
    canvas.element(ContentElement::ClipPath(ClipPathData {
        operations: rectangle(x - spill.0, y, w + spill.0 + spill.1, h),
        even_odd: false,
    }));
    for (i, text) in lines.iter().enumerate() {
        let width = canvas.width(text, family, size, style.bold, style.italic);
        // "general" is what LibreOffice (and some generators) write for
        // Excel's default: it aligns by the value's type, like no setting.
        let align =
            style
                .horizontal_align
                .as_deref()
                .filter(|a| *a != "general")
                .unwrap_or(if matches!(value, CellValue::Number(_)) {
                    "right"
                } else if matches!(value, CellValue::Boolean(_) | CellValue::Error(_)) {
                    "center"
                } else {
                    "left"
                });
        let at = match align {
            "right" => x + w - width - 2. * scale,
            "center" | "centerContinuous" => x + (w - width) / 2.,
            _ => x + 2. * scale + f64::from(style.indent) * size,
        };
        let baseline = top + size + i as f64 * line_h;
        let color = ink(style.font_color.as_deref());
        // 均等割り付け: the characters spread over the cell's width with equal
        // gaps, half a gap before the first and after the last, as Excel
        // prints it; a single character sits in the middle.
        let chars: Vec<char> = text.chars().collect();
        let room = w - 4. * scale - 2. * f64::from(style.indent) * size;
        let (at, width) = if align == "distributed" && chars.len() > 1 && width < room {
            let gap = (room - width) / chars.len() as f64;
            let mut pen = x + 2. * scale + f64::from(style.indent) * size + gap / 2.;
            for ch in &chars {
                let one = ch.to_string();
                canvas.text(&one, family, size, style.bold, style.italic, pen, baseline, color);
                pen += canvas.width(&one, family, size, style.bold, style.italic) + gap;
            }
            (x + 2. * scale + f64::from(style.indent) * size, room)
        } else {
            let at = if align == "distributed" { x + (w - width) / 2. } else { at };
            canvas.text(text, family, size, style.bold, style.italic, at, baseline, color);
            (at, width)
        };
        if style.underline {
            line(
                canvas,
                at,
                baseline + size * 0.12,
                at + width,
                baseline + size * 0.12,
                0.5 * scale,
                color,
            );
        }
        if style.strikethrough {
            line(
                canvas,
                at,
                baseline - size * 0.3,
                at + width,
                baseline - size * 0.3,
                0.5 * scale,
                color,
            );
        }
    }
    canvas.element(ContentElement::RestoreState);
}

/// Draw the print pages of every visible sheet of the workbook `data` onto
/// `canvas`: the print areas where the sheet sets them, the used range where
/// it does not. Returns what was left out.
pub fn pages(data: &[u8], canvas: &mut impl PrintCanvas) -> Result<Vec<Unsupported>, String> {
    let workbook = crate::parse_xlsx_preserving_values(data).map_err(|e| e.to_string())?;
    let mut archive = OoxmlArchive::new(data).map_err(|e| e.to_string())?;
    let parts = sheet_parts(&mut archive)?;
    let mut unsupported = Vec::new();
    let mut page_count = 0;
    for sheet in &workbook.sheets {
        if sheet.visibility != Visibility::Visible {
            continue;
        }
        let part = parts
            .get(&sheet.name)
            .ok_or_else(|| format!("missing sheet relationship for {}", sheet.name))?;
        let xml = archive.read_part(part).map_err(|e| e.to_string())?;
        let settings = Settings::read(&xml)?;
        let mut kinds = settings.unsupported.clone();
        kinds.extend(
            sheet
                .unsupported_elements
                .iter()
                .map(|v| format!("spreadsheet-{v}")),
        );
        if !sheet.drawings.is_empty() {
            kinds.insert("spreadsheet-drawings".into());
        }
        if !sheet.tables.is_empty() {
            kinds.insert("table-style".into());
        }
        if !sheet.conditional_rules.is_empty() {
            kinds.insert("conditional-formatting".into());
        }
        if !sheet.form_controls.is_empty() {
            kinds.insert("form-controls".into());
        }
        if sheet.header_footer.as_ref().is_some_and(|h| {
            [
                &h.odd_header,
                &h.odd_footer,
                &h.even_header,
                &h.even_footer,
                &h.first_header,
                &h.first_footer,
            ]
            .iter()
            .any(|v| !v.is_empty())
        }) {
            kinds.insert("header-footer".into());
        }
        for row in &sheet.rows {
            for c in &row.cells {
                if c.style.stacked_text || c.style.text_rotation.is_some_and(|v| v != 0) {
                    kinds.insert("rotated-cell-text".into());
                }
                if !c.runs.is_empty() {
                    kinds.insert("rich-cell-text".into());
                }
                if c.formula.is_some() && matches!(c.value, CellValue::Empty) {
                    kinds.insert("formula-without-cached-value".into());
                }
                for border in [
                    &c.style.border_top,
                    &c.style.border_bottom,
                    &c.style.border_left,
                    &c.style.border_right,
                ]
                .into_iter()
                .flatten()
                {
                    if !matches!(
                        border.style.as_str(),
                        "thin" | "medium" | "thick" | "hair" | "dotted" | "dashed" | "double"
                            | "mediumDashed" | "dashDot" | "mediumDashDot" | "dashDotDot"
                            | "mediumDashDotDot" | "slantDashDot"
                    ) {
                        kinds.insert("complex-cell-border".into());
                    }
                }
                if c.style.border_diagonal.is_some() {
                    kinds.insert("diagonal-cell-border".into());
                }
                if c.style
                    .pattern_fill
                    .as_ref()
                    .is_some_and(|f| !matches!(f.pattern.as_str(), "solid" | "none"))
                {
                    kinds.insert("pattern-cell-fill".into());
                }
                if c.style
                    .horizontal_align
                    .as_deref()
                    .is_some_and(|v| matches!(v, "justify" | "fill"))
                {
                    kinds.insert("complex-cell-alignment".into());
                }
            }
        }
        let mut print_areas = None;
        for (name, value) in &workbook.defined_names {
            if let Some((owner, leaf)) = crate::ir::defined_name_scope(name) {
                if owner == sheet.name {
                    let source = workbook
                        .defined_name_attributes
                        .get(name)
                        .and_then(|a| a.get("name"))
                        .map(String::as_str)
                        .unwrap_or(leaf);
                    if source == "_xlnm.Print_Area" {
                        print_areas = Some(areas(value)?);
                    }
                    if source == "_xlnm.Print_Titles" {
                        kinds.insert("print-titles".into());
                    }
                }
            }
        }
        let ranges = print_areas.unwrap_or_else(|| used(sheet).into_iter().collect());
        for area in ranges {
            page_count += render_area(canvas, sheet, &workbook, area, &settings, &mut kinds)?;
            if page_count > MAX_PAGES {
                return Err("spreadsheet exceeds the PDF page limit".into());
            }
        }
        for kind in kinds {
            unsupported.push(Unsupported {
                kind,
                part: part.clone(),
                count: 1,
            });
        }
    }
    if page_count == 0 {
        return Err("workbook has no visible printable cells".into());
    }
    Ok(unsupported)
}
/// Text that Excel lets run on over empty neighbours: strings that are
/// neither wrapped nor shrunk to fit. Numbers never do (they show ####).
fn overflows(value: &CellValue, style: &CellStyle) -> bool {
    matches!(value, CellValue::String(_) | CellValue::Utf16(_))
        && !style.wrap_text
        && !style.shrink_to_fit
}

/// The pixels of a column whose width the file states. A stored width
/// already carries the five-pixel gutter beside the text (ECMA-376 18.3.1.13:
/// 2.5 in a book whose digits are 8 pixels wide is a 20-pixel column), so it
/// is only scaled by the digit width, with Excel's rounding.
fn column_pixels(width: f64, digit: f64) -> f64 {
    (((256. * width + (128. / digit).trunc()) / 256.) * digit).trunc()
}

/// The look of a merged area: its top-left cell's, with each border taken from
/// the cells along that edge. Excel draws a merged area's bottom rule from its
/// bottom row and its right rule from its right column, so a box ruled all
/// round carries its right and bottom borders on cells other than the first.
fn merged_edges<'a>(
    first: &CellStyle,
    m: &MergeCell,
    style_at: impl Fn(u32, u32) -> Option<&'a CellStyle>,
) -> CellStyle {
    let mut out = first.clone();
    let along = |cells: Vec<(u32, u32)>, side: fn(&CellStyle) -> &Option<BorderLine>| {
        cells.into_iter().find_map(|(r, c)| style_at(r, c).and_then(|s| side(s).clone()))
    };
    let cols = || m.start_col..=m.end_col;
    let rows = || m.start_row..=m.end_row;
    out.border_top = along(cols().map(|c| (m.start_row, c)).collect(), |s| &s.border_top);
    out.border_bottom = along(cols().map(|c| (m.end_row, c)).collect(), |s| &s.border_bottom);
    out.border_left = along(rows().map(|r| (r, m.start_col)).collect(), |s| &s.border_left);
    out.border_right = along(rows().map(|r| (r, m.end_col)).collect(), |s| &s.border_right);
    out
}

/// A cell that overflowing text may cross: no value, no fill, not merged.
fn is_free(sheet: &Sheet, rows: &BTreeMap<u32, &crate::ir::Row>, row: u32, col: u32) -> bool {
    let merged = sheet.merge_cells.iter().any(|m| {
        row >= m.start_row && row <= m.end_row && col >= m.start_col && col <= m.end_col
    });
    let cell = rows.get(&row).and_then(|r| r.cells.iter().find(|c| c.col == col));
    !merged
        && cell.is_none_or(|c| {
            c.style.bg_color.is_none()
                && match &c.value {
                    CellValue::Empty => true,
                    CellValue::String(v) => v.is_empty(),
                    _ => false,
                }
        })
}

fn render_area(
    canvas: &mut impl PrintCanvas,
    sheet: &Sheet,
    book: &Workbook,
    area: Area,
    s: &Settings,
    kinds: &mut BTreeSet<String>,
) -> Result<usize, String> {
    let (c1, r1, c2, r2) = area;
    if u64::from(c2 - c1 + 1) * u64::from(r2 - r1 + 1) > MAX_CELLS {
        return Err("print area exceeds the PDF cell limit".into());
    }
    let family = book.default_style.font_name.as_deref().unwrap_or("Calibri");
    let digit = canvas.width(
        "0",
        family,
        f64::from(book.default_style.font_size.unwrap_or(11.)),
        false,
        false,
    ) / 0.75;
    // Excel counts the digit width in whole pixels.
    let digit = digit.round().max(1.);
    let columns: Vec<f64> = (c1..=c2)
        .map(|c| {
            if sheet.hidden_cols.contains(&c) {
                return 0.;
            }
            let stored = sheet
                .col_widths
                .get(c as usize)
                .copied()
                .filter(|w| *w > 0.)
                .or((sheet.default_col_width > 0.).then_some(sheet.default_col_width));
            match stored {
                Some(width) => column_pixels(f64::from(width), digit) * 0.75,
                // No width anywhere: Excel's 8.43 characters and the gutter.
                None => (8.43 * digit + 5.).floor() * 0.75,
            }
        })
        .collect();
    let rows: BTreeMap<u32, _> = sheet.rows.iter().map(|r| (r.index, r)).collect();
    let default_height = if sheet.default_row_height > 0. {
        sheet.default_row_height
    } else {
        15.
    };
    let heights: Vec<f64> = (r1..=r2)
        .map(|r| {
            rows.get(&r).map_or(f64::from(default_height), |row| {
                if row.hidden {
                    0.
                } else {
                    f64::from(row.height.unwrap_or(default_height)).max(0.)
                }
            })
        })
        .collect();
    if columns.iter().sum::<f64>() <= 0. || heights.iter().sum::<f64>() <= 0. {
        return Ok(0);
    }
    let capacity_x = s.width - s.left - s.right;
    let capacity_y = s.height - s.top - s.bottom;
    let mut scale = s.scale;
    if s.fit && (s.fit_w > 0 || s.fit_h > 0) {
        let mut low = 0.000001;
        let mut high = 1.;
        let empty = BTreeSet::new();
        for _ in 0..45 {
            let mid = (low + high) / 2.;
            let x = axis(&columns, capacity_x / mid, c1, &empty).len();
            let y = axis(&heights, capacity_y / mid, r1, &empty).len();
            let widest = columns.iter().copied().fold(0., f64::max) * mid;
            let tallest = heights.iter().copied().fold(0., f64::max) * mid;
            if (s.fit_w == 0 || x <= s.fit_w)
                && (s.fit_h == 0 || y <= s.fit_h)
                && widest <= capacity_x
                && tallest <= capacity_y
            {
                low = mid;
            } else {
                high = mid;
            }
        }
        scale = low;
    }
    let empty = BTreeSet::new();
    // Row break id=N is after row N; columns have zero-based coordinates here.
    let row_breaks: BTreeSet<u32> = s.row_breaks.iter().map(|r| r + 1).collect();
    let xs = axis(
        &columns,
        capacity_x / scale,
        c1,
        if s.fit { &empty } else { &s.col_breaks },
    );
    let ys = axis(
        &heights,
        capacity_y / scale,
        r1,
        if s.fit { &empty } else { &row_breaks },
    );
    if xs.len().saturating_mul(ys.len()) > MAX_PAGES {
        return Err("spreadsheet exceeds the PDF page limit".into());
    }
    let cx = positions(&columns);
    let cy = positions(&heights);
    let mut tiles = Vec::new();
    for m in &sheet.merge_cells {
        if m.end_col < c1 || m.start_col > c2 || m.end_row < r1 || m.start_row > r2 {
            continue;
        }
        if m.start_col < c1
            || m.end_col > c2
            || m.start_row < r1
            || m.end_row > r2
            || xs
                .iter()
                .skip(1)
                .any(|(at, _)| m.start_col < c1 + *at as u32 && m.end_col >= c1 + *at as u32)
            || ys
                .iter()
                .skip(1)
                .any(|(at, _)| m.start_row < r1 + *at as u32 && m.end_row >= r1 + *at as u32)
        {
            kinds.insert("merge-across-print-boundary".into());
        }
    }
    if s.over {
        for &y in &ys {
            for &x in &xs {
                tiles.push((x, y));
            }
        }
    } else {
        for &x in &xs {
            for &y in &ys {
                tiles.push((x, y));
            }
        }
    }
    for &((x1, x2), (y1, y2)) in &tiles {
        canvas.page(s.width, s.height);
        let offset_x = s.left
            + if s.center_x {
                (capacity_x - (cx[x2] - cx[x1]) * scale) / 2.
            } else {
                0.
            };
        let offset_y = s.top
            + if s.center_y {
                (capacity_y - (cy[y2] - cy[y1]) * scale) / 2.
            } else {
                0.
            };
        canvas.element(ContentElement::SaveState);
        canvas.element(ContentElement::ClipPath(ClipPathData {
            operations: rectangle(s.left, s.top, capacity_x, capacity_y),
            even_odd: false,
        }));
        for ri in y1..y2 {
            if heights[ri] <= 0. {
                continue;
            }
            for ci in x1..x2 {
                if columns[ci] <= 0. {
                    continue;
                }
                let row = r1 + ri as u32;
                let col = c1 + ci as u32;
                let merge = sheet.merge_cells.iter().find(|m| {
                    row >= m.start_row && row <= m.end_row && col >= m.start_col && col <= m.end_col
                });
                let (mut end_col, mut end_row) = (ci + 1, ri + 1);
                if let Some(m) = merge {
                    if row != m.start_row || col != m.start_col {
                        continue;
                    }
                    end_col = ((m.end_col - c1 + 1) as usize).min(columns.len());
                    end_row = ((m.end_row - r1 + 1) as usize).min(heights.len());
                }
                let value = rows
                    .get(&row)
                    .and_then(|r| r.cells.iter().find(|c| c.col == col));
                let default = CellStyle::default();
                let mut style = value.map_or(&default, |c| &c.style);
                let edged;
                if let Some(m) = merge {
                    edged = merged_edges(style, m, |r, c| {
                        rows.get(&r).and_then(|x| x.cells.iter().find(|x| x.col == c)).map(|x| &x.style)
                    });
                    style = &edged;
                }
                let blank = CellValue::Empty;
                let value = value.map_or(&blank, |c| &c.value);
                let spill = if merge.is_none() && overflows(value, style) {
                    let free = |c: usize| {
                        columns[c] <= 0. || is_free(sheet, &rows, row, c1 + c as u32)
                    };
                    let mut left = ci;
                    while left > x1 && free(left - 1) {
                        left -= 1;
                    }
                    let mut right = ci + 1;
                    while right < x2 && free(right) {
                        right += 1;
                    }
                    let (l, r) = ((cx[ci] - cx[left]) * scale, (cx[right] - cx[ci + 1]) * scale);
                    // Text runs on in the direction it is aligned to grow.
                    match style.horizontal_align.as_deref() {
                        None | Some("general") | Some("left") => (0., r),
                        Some("right") => (l, 0.),
                        Some("center") | Some("centerContinuous") => (l, r),
                        _ => (0., 0.),
                    }
                } else {
                    (0., 0.)
                };
                cell(
                    canvas,
                    value,
                    style,
                    offset_x + (cx[ci] - cx[x1]) * scale,
                    offset_y + (cy[ri] - cy[y1]) * scale,
                    (cx[end_col] - cx[ci]) * scale,
                    (cy[end_row] - cy[ri]) * scale,
                    scale,
                    s.grid,
                    spill,
                    s.mono,
                );
            }
        }
        canvas.element(ContentElement::RestoreState);
    }
    Ok(tiles.len())
}

#[cfg(test)]
mod tests {
    use super::*;
    #[test]
    fn a_stored_column_width_already_carries_its_gutter() {
        // Calibri 11 (7-pixel digits): Excel's default 8.43 characters is
        // stored as 9.140625 and drawn 64 pixels wide.
        assert_eq!(column_pixels(9.140625, 7.), 64.);
        // Yu Gothic 11 (8-pixel digits): 2.5 is a 20-pixel square cell.
        assert_eq!(column_pixels(2.5, 8.), 20.);
        // MS PGothic 11 (9-pixel digits), as a city workbook stores it.
        assert_eq!(column_pixels(9.44140625, 9.), 85.);
    }

    #[test]
    fn a_merged_area_takes_each_border_from_the_cells_on_that_edge() {
        let line = || Some(serde_json::from_str::<BorderLine>(r#"{"style":"thin"}"#).unwrap());
        let first = CellStyle { border_top: line(), border_left: line(), ..Default::default() };
        let right = CellStyle { border_top: line(), border_right: line(), ..Default::default() };
        let bottom = CellStyle { border_bottom: line(), border_left: line(), ..Default::default() };
        let m: MergeCell = serde_json::from_str(r#"{"start_row":1,"end_row":2,"start_col":0,"end_col":1}"#).unwrap();
        let edged = merged_edges(&first, &m, |r, c| match (r, c) {
            (1, 0) => Some(&first),
            (1, 1) => Some(&right),
            (2, 0) => Some(&bottom),
            _ => None,
        });
        assert!(edged.border_top.is_some() && edged.border_left.is_some());
        assert!(edged.border_right.is_some(), "the right column carries the right rule");
        assert!(edged.border_bottom.is_some(), "the bottom row carries the bottom rule");
        let alone = merged_edges(&first, &m, |r, c| (r, c).eq(&(1, 0)).then_some(&first));
        assert!(alone.border_right.is_none() && alone.border_bottom.is_none());
    }

    #[test]
    fn print_areas_and_margins_are_validated() {
        assert!(areas("'A,B'!$D$2:$A$1").is_err());
        assert_eq!(
            areas("'A,B'!$A$1:$B$2,'A,B'!$D$3").unwrap(),
            vec![(0, 1, 1, 2), (3, 3, 3, 3)]
        );
        assert!(Settings::read(r#"<worksheet><pageMargins left="100"/></worksheet>"#).is_err());
    }
}
