// SPDX-License-Identifier: MIT OR Apache-2.0

//! Where each `{{placeholder}}` of a Word template lands on the printed page.
//!
//! A form designed once in Word can be filled on screen, printed blank and
//! filled by hand, or scanned back: reading a scan needs to know which area of
//! the page holds which field. The layout already places every character, so
//! the placeholders' boxes come from the same layout that prints the form,
//! together with the table cell around each (the area a person writes in).

use oxidocs_core::layout::{LayoutContent, LayoutPage};
use serde::Serialize;
use wasm_bindgen::prelude::*;

#[derive(Serialize, Debug, Clone, Copy, PartialEq)]
struct Rect {
    x: f32,
    y: f32,
    width: f32,
    height: f32,
}

#[derive(Serialize, Debug)]
struct Field {
    /// The text between the braces, trimmed: `氏名`, `契約日:和暦`, `明細.品名`.
    name: String,
    /// 1-based page number.
    page: usize,
    /// The placeholder's own text, one box per line it occupies.
    text: Vec<Rect>,
    /// The innermost table cell around the placeholder, when it is in one.
    cell: Option<Rect>,
}

#[derive(Serialize)]
struct Pages {
    /// Page sizes in points, in order.
    pages: Vec<(f32, f32)>,
    fields: Vec<Field>,
}

/// The placeholders of a .docx template with their positions on the laid-out
/// pages, as JSON: `{"pages": [[width, height], …], "fields": [{"name",
/// "page", "text": [{x, y, width, height}…], "cell": {…} | null}]}`, in points
/// from the top-left corner of the page.
#[wasm_bindgen]
pub fn docx_field_boxes(data: &[u8]) -> Result<String, JsError> {
    let doc = oxidocs_core::parse_docx(data).map_err(|e| JsError::new(&e.to_string()))?;
    let layout = oxidocs_core::layout::LayoutEngine::for_document(&doc).layout(&doc);
    let out = Pages {
        pages: layout.pages.iter().map(|p| (p.width, p.height)).collect(),
        fields: layout.pages.iter().enumerate().flat_map(|(i, p)| page_fields(p, i + 1)).collect(),
    };
    serde_json::to_string(&out).map_err(|e| JsError::new(&e.to_string()))
}

fn page_fields(page: &LayoutPage, number: usize) -> Vec<Field> {
    // The page's text in drawing order, each character mapped to the box of
    // the fragment that draws it.
    let mut chars: Vec<(char, Rect)> = Vec::new();
    for e in &page.elements {
        if let LayoutContent::Text { text, .. } = &e.content {
            let n = text.chars().count().max(1) as f32;
            for (k, ch) in text.chars().enumerate() {
                // A fragment's characters share its box, split evenly; close
                // enough to tell lines apart and to bound the placeholder.
                let w = e.width / n;
                chars.push((ch, Rect { x: e.x + w * k as f32, y: e.y, width: w, height: e.height }));
            }
        }
    }
    let borders = borders(page);
    let mut fields = Vec::new();
    let mut i = 0;
    while i + 1 < chars.len() {
        if chars[i].0 == '{' && chars[i + 1].0 == '{' {
            if let Some(end) = (i + 2..chars.len().saturating_sub(1)).find(|&j| chars[j].0 == '}' && chars[j + 1].0 == '}') {
                let name: String = chars[i + 2..end].iter().map(|c| c.0).collect();
                let name = name.trim().to_string();
                if !name.is_empty() && !name.contains('{') {
                    let text = lines(&chars[i..end + 2]);
                    let cell = text.first().and_then(|r| enclosing_cell(&borders, r));
                    fields.push(Field { name, page: number, text, cell });
                    i = end + 2;
                    continue;
                }
            }
        }
        i += 1;
    }
    fields
}

/// The boxes of a run of characters, merged per line.
fn lines(chars: &[(char, Rect)]) -> Vec<Rect> {
    let mut out: Vec<Rect> = Vec::new();
    for (_, r) in chars {
        match out.last_mut() {
            Some(last) if (last.y - r.y).abs() < r.height.max(1.) * 0.5 => {
                let right = (last.x + last.width).max(r.x + r.width);
                let bottom = (last.y + last.height).max(r.y + r.height);
                last.x = last.x.min(r.x);
                last.y = last.y.min(r.y);
                last.width = right - last.x;
                last.height = bottom - last.y;
            }
            _ => out.push(*r),
        }
    }
    out
}

/// Table border segments: (horizontal, fixed coordinate, from, to).
fn borders(page: &LayoutPage) -> Vec<(bool, f32, f32, f32)> {
    page.elements
        .iter()
        .filter_map(|e| match e.content {
            LayoutContent::TableBorder { x1, y1, x2, y2, .. } => {
                if (y1 - y2).abs() < 0.5 {
                    Some((true, y1, x1.min(x2), x1.max(x2)))
                } else if (x1 - x2).abs() < 0.5 {
                    Some((false, x1, y1.min(y2), y1.max(y2)))
                } else {
                    None
                }
            }
            _ => None,
        })
        .collect()
}

/// The smallest box of table borders around the middle of `r`: the nearest
/// border on each side whose extent covers that point.
fn enclosing_cell(borders: &[(bool, f32, f32, f32)], r: &Rect) -> Option<Rect> {
    let (cx, cy) = (r.x + r.width / 2., r.y + r.height / 2.);
    let nearest = |horizontal: bool, along: f32, before: bool, from: f32| {
        borders
            .iter()
            .filter(|b| b.0 == horizontal && b.2 - 0.5 <= along && along <= b.3 + 0.5)
            .map(|b| b.1)
            .filter(|&at| if before { at <= from } else { at >= from })
            .min_by(|a, b| (a - from).abs().total_cmp(&(b - from).abs()))
    };
    let top = nearest(true, cx, true, cy)?;
    let bottom = nearest(true, cx, false, cy)?;
    let left = nearest(false, cy, true, cx)?;
    let right = nearest(false, cy, false, cx)?;
    Some(Rect { x: left, y: top, width: right - left, height: bottom - top })
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn a_placeholder_split_over_lines_gets_a_box_per_line() {
        let a = Rect { x: 10., y: 0., width: 5., height: 10. };
        let b = Rect { x: 15., y: 0., width: 5., height: 10. };
        let c = Rect { x: 0., y: 12., width: 5., height: 10. };
        let merged = lines(&[('{', a), ('{', b), ('x', c)]);
        assert_eq!(merged.len(), 2);
        assert_eq!(merged[0], Rect { x: 10., y: 0., width: 10., height: 10. });
    }

    #[test]
    fn the_cell_is_the_nearest_border_on_each_side() {
        let borders = vec![
            (true, 0., 0., 100.), (true, 20., 0., 100.), (true, 40., 0., 100.),
            (false, 0., 0., 40.), (false, 50., 0., 40.), (false, 100., 0., 40.),
        ];
        let text = Rect { x: 55., y: 25., width: 10., height: 8. };
        assert_eq!(enclosing_cell(&borders, &text), Some(Rect { x: 50., y: 20., width: 50., height: 20. }));
        let outside = Rect { x: 120., y: 25., width: 10., height: 8. };
        assert_eq!(enclosing_cell(&borders, &outside), None);
    }
}
