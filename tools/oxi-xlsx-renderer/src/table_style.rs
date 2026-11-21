// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! Excel's built-in table styles, laid onto the cells of each table.
//!
//! `table_styles.json` is what Excel itself holds for its sixty built-in
//! styles (`tools/metrics/_xlsx_table_styles_dump.py`, read element by
//! element through `TableStyles(..).TableStyleElements(..)`): for each
//! element, its fill, font and six borders, colours as a theme slot plus
//! TintAndShade. The elements stack in ECMA-376's order -- whole table,
//! column stripes, row stripes, last column, first column, header row,
//! total row, then the four corner cells -- and the highest one that says
//! something about a property decides it. A cell's own format wins over the
//! table's wherever the cell states one.

use oxicells_core::ir::{BorderLine, CellStyle, CellValue, Sheet, Table};
use serde_json::Value;

static DATA: &str = include_str!("table_styles.json");

fn styles() -> &'static Value {
    use std::sync::OnceLock;
    static PARSED: OnceLock<Value> = OnceLock::new();
    PARSED.get_or_init(|| serde_json::from_str(DATA).unwrap_or(Value::Null))
}

/// Whether the style is one of the sixty this module knows.
pub(crate) fn known(table: &Table) -> bool {
    table
        .style
        .as_deref()
        .is_some_and(|name| styles().get(name).is_some())
}

/// A colour as Excel holds it in a table style: theme slot + tint, read
/// against the workbook's own theme. COM numbers the slots 1 = lt1, 2 = dk1,
/// 5..10 = accent1..6 (the sixty styles use no other).
fn colour(spec: &Value, theme: &[String]) -> Option<String> {
    if let Some(slot) = spec.get("theme").and_then(Value::as_u64) {
        let index = match slot {
            1 => 1,
            2 => 0,
            3 => 3,
            4 => 2,
            k @ 5..=12 => k as usize - 1,
            _ => return spec.get("rgb").and_then(Value::as_str).map(str::to_string),
        };
        let base = theme.get(index)?;
        let tint = spec.get("tint").and_then(Value::as_f64).unwrap_or(0.0);
        return Some(oxicells_core::parser::tint_cell_color(base, tint));
    }
    spec.get("rgb").and_then(Value::as_str).map(str::to_string)
}

/// COM LineStyle + Weight -> the OOXML border name.
fn border_name(style: i64, weight: i64) -> &'static str {
    match (style, weight) {
        (-4119, _) => "double",
        (-4115, _) => "dashed",
        (-4118, _) => "dotted",
        (4, _) => "dashDot",
        (5, _) => "dashDotDot",
        (_, 1) => "hair",
        (_, -4138) => "medium",
        (_, 4) => "thick",
        _ => "thin",
    }
}

fn border(spec: &Value, theme: &[String]) -> Option<BorderLine> {
    let style = spec.get("style").and_then(Value::as_i64)?;
    let weight = spec.get("weight").and_then(Value::as_i64).unwrap_or(2);
    Some(BorderLine {
        style: border_name(style, weight).to_string(),
        color: colour(spec, theme),
        color_reference: None,
    })
}

#[derive(Default, Clone)]
struct Layer {
    fill: Option<String>,
    font_colour: Option<String>,
    bold: Option<bool>,
    italic: Option<bool>,
    top: Option<BorderLine>,
    bottom: Option<BorderLine>,
    left: Option<BorderLine>,
    right: Option<BorderLine>,
}

/// One element laid over a rectangle `(r0, c0, r1, c1)`: its outer borders
/// on the rectangle's edges, its inside borders between its cells.
fn lay(
    out: &mut Layer,
    element: &Value,
    theme: &[String],
    (row, column): (u32, u32),
    (r0, c0, r1, c1): (u32, u32, u32, u32),
) {
    if let Some(fill) = element.get("fill") {
        if let Some(c) = colour(fill, theme) {
            out.fill = Some(c);
        }
    }
    if let Some(font) = element.get("font") {
        if let Some(c) = font.get("colour").and_then(|c| colour(c, theme)) {
            out.font_colour = Some(c);
        }
        if font.get("bold").and_then(Value::as_bool) == Some(true) {
            out.bold = Some(true);
        }
        if font.get("italic").and_then(Value::as_bool) == Some(true) {
            out.italic = Some(true);
        }
    }
    let Some(b) = element.get("borders") else {
        return;
    };
    let pick = |outer: &str, inner: &str, on_edge: bool| {
        b.get(if on_edge { outer } else { inner })
            .and_then(|s| border(s, theme))
    };
    if let Some(line) = pick("top", "insideH", row == r0) {
        out.top = Some(line);
    }
    if let Some(line) = pick("bottom", "insideH", row == r1) {
        out.bottom = Some(line);
    }
    if let Some(line) = pick("left", "insideV", column == c0) {
        out.left = Some(line);
    }
    if let Some(line) = pick("right", "insideV", column == c1) {
        out.right = Some(line);
    }
}

fn layer_for(table: &Table, theme: &[String], row: u32, column: u32) -> Option<Layer> {
    let style = styles().get(table.style.as_deref()?)?;
    let (top, bottom, left, right) = (
        table.start_row,
        table.end_row,
        table.start_col,
        table.end_col,
    );
    let header_end = top + table.header_rows; // first data row
    let total_start = bottom + 1 - table.totals_rows.min(bottom + 1 - top); // first total row
    let in_header = row < header_end;
    let in_total = table.totals_rows > 0 && row >= total_start;
    let mut out = Layer::default();
    let el = |name: &str| style.get(name);
    let at = (row, column);
    if let Some(e) = el("wholeTable") {
        lay(&mut out, e, theme, at, (top, left, bottom, right));
    }
    let (data_top, data_bottom) = (header_end, total_start.saturating_sub(1));
    if table.banded_columns && !in_header && !in_total {
        let name = if (column - left) % 2 == 0 {
            "firstColumnStripe"
        } else {
            "secondColumnStripe"
        };
        if let Some(e) = el(name) {
            lay(&mut out, e, theme, at, (data_top, column, data_bottom, column));
        }
    }
    if table.banded_rows && !in_header && !in_total {
        let name = if (row - data_top) % 2 == 0 {
            "firstRowStripe"
        } else {
            "secondRowStripe"
        };
        if let Some(e) = el(name) {
            lay(&mut out, e, theme, at, (row, left, row, right));
        }
    }
    if table.emphasize_last_column && column == right {
        if let Some(e) = el("lastColumn") {
            lay(&mut out, e, theme, at, (top, right, bottom, right));
        }
    }
    if table.emphasize_first_column && column == left {
        if let Some(e) = el("firstColumn") {
            lay(&mut out, e, theme, at, (top, left, bottom, left));
        }
    }
    if in_header {
        let last = header_end - 1;
        if let Some(e) = el("headerRow") {
            lay(&mut out, e, theme, at, (top, left, last, right));
        }
        if column == left {
            if let Some(e) = el("firstHeaderCell") {
                lay(&mut out, e, theme, at, (top, left, last, left));
            }
        }
        if column == right {
            if let Some(e) = el("lastHeaderCell") {
                lay(&mut out, e, theme, at, (top, right, last, right));
            }
        }
    }
    if in_total {
        if let Some(e) = el("totalRow") {
            lay(&mut out, e, theme, at, (total_start, left, bottom, right));
        }
        if column == left {
            if let Some(e) = el("firstTotalCell") {
                lay(&mut out, e, theme, at, (total_start, left, bottom, left));
            }
        }
        if column == right {
            if let Some(e) = el("lastTotalCell") {
                lay(&mut out, e, theme, at, (total_start, right, bottom, right));
            }
        }
    }
    Some(out)
}

/// Whether a cell's font is the workbook's Normal font, so the table may
/// recolour and embolden it. A cell that names a font of its own keeps it.
fn wears_normal_font(style: &CellStyle) -> bool {
    style.normal_font_link != Some(false)
}

fn apply(style: &mut CellStyle, layer: &Layer) {
    if style.bg_color.is_none() && style.pattern_fill.is_none() {
        style.bg_color.clone_from(&layer.fill);
    }
    if wears_normal_font(style) {
        if let Some(c) = &layer.font_colour {
            style.font_color = Some(c.clone());
        }
        if layer.bold == Some(true) {
            style.bold = true;
        }
        if layer.italic == Some(true) {
            style.italic = true;
        }
    }
    for (own, theirs) in [
        (&mut style.border_top, &layer.top),
        (&mut style.border_bottom, &layer.bottom),
        (&mut style.border_left, &layer.left),
        (&mut style.border_right, &layer.right),
    ] {
        if own.is_none() {
            own.clone_from(theirs);
        }
    }
}

/// The sheet with every built-in table style laid onto its cells, empty
/// ones written out where the style shows on them.
pub(crate) fn with_table_styles(sheet: &Sheet, theme: &[String]) -> Sheet {
    let mut out = sheet.clone();
    for table in sheet.tables.iter().filter(|t| known(t)) {
        for row_index in table.start_row..=table.end_row {
            let at = match out.rows.binary_search_by_key(&row_index, |row| row.index) {
                Ok(at) => at,
                Err(at) => {
                    out.rows.insert(
                        at,
                        oxicells_core::ir::Row {
                            index: row_index,
                            cells: Vec::new(),
                            height: None,
                            custom_height: false,
                            format_recorded: None,
                            style_font: None,
                            thick_top: false,
                            thick_bottom: false,
                            hidden: false,
                        },
                    );
                    at
                }
            };
            let row = &mut out.rows[at];
            for column in table.start_col..=table.end_col {
                let Some(layer) = layer_for(table, theme, row_index, column) else {
                    continue;
                };
                match row.cells.iter_mut().find(|cell| cell.col == column) {
                    Some(cell) => apply(&mut cell.style, &layer),
                    None => {
                        let mut style = CellStyle::default();
                        apply(&mut style, &layer);
                        row.cells.push(oxicells_core::ir::Cell {
                            col: column,
                            value: CellValue::Empty,
                            style,
                            formula: None,
                            formula_utf16: None,
                            runs: Vec::new(),
                            array_block: None,
                            spill: Default::default(),
                        });
                    }
                }
            }
            row.cells.sort_by_key(|cell| cell.col);
        }
    }
    out
}

// ---- pivot tables ------------------------------------------------------------

static PIVOT_DATA: &str = include_str!("pivot_styles.json");

fn pivot_styles() -> &'static Value {
    use std::sync::OnceLock;
    static PARSED: OnceLock<Value> = OnceLock::new();
    PARSED.get_or_init(|| serde_json::from_str(PIVOT_DATA).unwrap_or(Value::Null))
}

/// The sheet with each pivot table's built-in style laid onto its cells.
///
/// The parts a pivot style dresses are the pivot's own -- heading rows, the
/// rows of an outer item (a row subheading on its label, a subtotal on its
/// numbers, since the compact form puts an item's total on its own row),
/// subtotal and grand total rows, the grand total column, page fields -- so
/// the grid is worked out again (`pivot::compute`) to know which row is
/// which. Elements stack low to high as Excel stacks them; a cell's own
/// format wins.
pub(crate) fn with_pivot_styles(
    sheet: &Sheet,
    workbook: &oxicells_core::ir::Workbook,
    bytes: &[u8],
    theme: &[String],
) -> Sheet {
    use oxicells_core::pivot::{self, RowKind};
    let mut out = sheet.clone();
    let specs = pivot::read_pivots(bytes);
    for spec in specs.iter().filter(|p| p.sheet.eq_ignore_ascii_case(&sheet.name)) {
        let Some(style) = spec.style.as_deref().and_then(|name| pivot_styles().get(name)) else {
            continue;
        };
        let get = pivot::source_reader(workbook, &spec.source_sheet);
        let grid = pivot::compute(spec, &get, pivot::Ui::Japanese);
        let width = grid.cells.iter().map(Vec::len).max().unwrap_or(0);
        if width == 0 {
            continue;
        }
        let lc = grid.label_columns.min(width);
        let rows = grid.cells.len();
        let first_table_row = grid
            .kinds
            .iter()
            .position(|k| !matches!(k, RowKind::Page | RowKind::Blank))
            .unwrap_or(0);
        let header_end = grid
            .kinds
            .iter()
            .rposition(|k| *k == RowKind::Header)
            .map(|i| i + 1)
            .unwrap_or(first_table_row);
        let grand_col = spec.cols.iter().any(Option::is_some) && spec.col_grand_totals;
        let last_row = rows as u32 - 1;
        let last_col = width as u32 - 1;
        for (i, kind) in grid.kinds.iter().enumerate() {
            let row_index = grid.origin.0 + i as u32;
            let at = match out.rows.binary_search_by_key(&(row_index + 1), |row| row.index) {
                Ok(at) => at,
                Err(at) => {
                    out.rows.insert(
                        at,
                        oxicells_core::ir::Row {
                            index: row_index + 1,
                            cells: Vec::new(),
                            height: None,
                            custom_height: false,
                            format_recorded: None,
                            style_font: None,
                            thick_top: false,
                            thick_bottom: false,
                            hidden: false,
                        },
                    );
                    at
                }
            };
            // A table of page fields alone has no body to dress: blind
            // 41601's seven `Name (All)` tables show their page row and
            // nothing under it.
            let bare = spec.rows.is_empty() && spec.cols.is_empty() && spec.data.is_empty();
            let cols = match kind {
                RowKind::Page => 2,
                RowKind::Blank => 0,
                _ if bare => 0,
                _ => width,
            };
            let r = i as u32;
            for j in 0..cols {
                let c = j as u32;
                let col = grid.origin.1 + c;
                let mut layer = Layer::default();
                let el = |name: &str| style.get(name);
                let here = (r, c);
                if *kind == RowKind::Page {
                    let name = if j == 0 { "pageFieldLabels" } else { "pageFieldValues" };
                    if let Some(e) = el(name) {
                        lay(&mut layer, e, theme, here, (r, c, r, c));
                    }
                } else {
                    if let Some(e) = el("wholeTable") {
                        lay(&mut layer, e, theme, here, (first_table_row as u32, 0, last_row, last_col));
                    }
                    let body = i >= header_end;
                    if body && spec.show_row_headers && j < lc {
                        if let Some(e) = el("firstColumn") {
                            lay(&mut layer, e, theme, here, (header_end as u32, 0, last_row, lc as u32 - 1));
                        }
                    }
                    if *kind == RowKind::Header && spec.show_col_headers {
                        if let Some(e) = el("headerRow") {
                            lay(&mut layer, e, theme, here, (first_table_row as u32, 0, header_end as u32 - 1, last_col));
                        }
                        if i == first_table_row && j == 0 {
                            if let Some(e) = el("firstHeaderCell") {
                                lay(&mut layer, e, theme, here, (r, 0, r, 0));
                            }
                        }
                    }
                    let line = (r, 0, r, last_col);
                    match kind {
                        RowKind::Item { level, leaf: false } => {
                            // The item's own row is a row subheading from edge
                            // to edge, its numbers included: under PivotStyleLight16
                            // the line beneath East is the heading's pale blue
                            // across both columns, not the subtotal's dark one.
                            let n = (level + 1).min(3);
                            if let Some(e) = el(&format!("rowSubheading{n}")) {
                                lay(&mut layer, e, theme, here, line);
                            }
                        }
                        RowKind::Subtotal { level } => {
                            if let Some(e) = el(&format!("subtotalRow{}", (level + 1).min(3))) {
                                lay(&mut layer, e, theme, here, line);
                            }
                        }
                        _ => {}
                    }
                    if grand_col && body && c == last_col {
                        if let Some(e) = el("grandTotalColumn") {
                            lay(&mut layer, e, theme, here, (header_end as u32, c, last_row, c));
                        }
                    }
                    if *kind == RowKind::Grand {
                        if let Some(e) = el("grandTotalRow") {
                            lay(&mut layer, e, theme, here, line);
                        }
                    }
                }
                let row = &mut out.rows[at];
                // A page field's button is written on its name and drawn on the
                // item beside it (p08: `pivotButton` on A1, the arrow in B1).
                let page_button = (*kind == RowKind::Page).then_some(j == 1);
                match row.cells.iter_mut().find(|cell| cell.col == col) {
                    Some(cell) => {
                        apply(&mut cell.style, &layer);
                        if let Some(button) = page_button {
                            cell.style.pivot_button = button;
                        }
                    }
                    None => {
                        let mut dressed = CellStyle::default();
                        apply(&mut dressed, &layer);
                        dressed.pivot_button = page_button.unwrap_or(false);
                        row.cells.push(oxicells_core::ir::Cell {
                            col,
                            value: CellValue::Empty,
                            style: dressed,
                            formula: None,
                            formula_utf16: None,
                            runs: Vec::new(),
                            array_block: None,
                            spill: Default::default(),
                        });
                    }
                }
            }
            out.rows[at].cells.sort_by_key(|cell| cell.col);
        }
    }
    out
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn theme_slots_and_tints_resolve_like_excel() {
        let theme: Vec<String> = [
            "000000", "FFFFFF", "0E2841", "E8E8E8", "156082", "E97132", "196B24", "0F9ED5",
            "A02B93", "4EA72E",
        ]
        .iter()
        .map(|s| s.to_string())
        .collect();
        let spec: Value = serde_json::json!({"theme": 5, "tint": 0.399976});
        assert_eq!(colour(&spec, &theme).as_deref(), Some("44B3E1"));
        let spec: Value = serde_json::json!({"theme": 1, "tint": -0.149998});
        assert_eq!(colour(&spec, &theme).as_deref(), Some("D9D9D9"));
        let spec: Value = serde_json::json!({"theme": 10, "tint": 0.799982});
        assert_eq!(colour(&spec, &theme).as_deref(), Some("DAF2D0"));
    }
}
