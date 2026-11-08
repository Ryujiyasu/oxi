// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! Reading for DrawingML parts drawn as groups (SmartArt drawings, charts):
//! a small element tree, theme colours with their modifiers, and the group
//! the drawn shapes and text boxes go into.

use quick_xml::events::Event;
use quick_xml::Reader;

use super::theme::ThemeColors;
use crate::ir::{Block, FloatingPosition, Run, RunStyle, TextBox, VectorShape, WrapType};

pub(super) const EMU_PER_PT: f32 = 12700.0;


/// The shapes and text boxes of a drawn group (a SmartArt, a chart),
/// positioned relative to its frame's top-left in points.
pub(super) struct Group {
    pub(super) shapes: Vec<VectorShape>,
    pub(super) texts: Vec<TextBox>,
}

/// One element of the drawing part, with the byte span it came from (a colour
/// element's modifiers are read again from its source by the shared colour
/// parser).
pub(super) struct Node {
    pub(super) name: String,
    pub(super) attrs: Vec<(String, String)>,
    pub(super) children: Vec<Node>,
    pub(super) text: String,
    span: (usize, usize),
}

impl Node {
    pub(super) fn attr(&self, key: &str) -> Option<&str> {
        self.attrs.iter().find(|(k, _)| k == key).map(|(_, v)| v.as_str())
    }
    pub(super) fn child(&self, name: &str) -> Option<&Node> {
        self.children.iter().find(|c| c.name == name)
    }
    pub(super) fn path(&self, names: &[&str]) -> Option<&Node> {
        names.iter().try_fold(self, |n, name| n.child(name))
    }
    pub(super) fn num(&self, key: &str) -> Option<f32> {
        self.attr(key)?.parse().ok()
    }
}

fn local(name: &[u8]) -> String {
    let s = String::from_utf8_lossy(name);
    s.rsplit(':').next().unwrap_or_default().to_string()
}

pub(super) fn parse_tree(xml: &str) -> Option<Node> {
    let mut reader = Reader::from_str(xml);
    let mut stack: Vec<Node> = vec![Node { name: String::new(), attrs: vec![], children: vec![], text: String::new(), span: (0, 0) }];
    loop {
        let start = reader.buffer_position() as usize;
        match reader.read_event().ok()? {
            Event::Start(e) => {
                let attrs = e
                    .attributes()
                    .flatten()
                    .map(|a| (local(a.key.as_ref()), String::from_utf8_lossy(&a.value).into_owned()))
                    .collect();
                stack.push(Node { name: local(e.name().as_ref()), attrs, children: vec![], text: String::new(), span: (start, 0) });
            }
            Event::Empty(e) => {
                let attrs = e
                    .attributes()
                    .flatten()
                    .map(|a| (local(a.key.as_ref()), String::from_utf8_lossy(&a.value).into_owned()))
                    .collect();
                let end = reader.buffer_position() as usize;
                stack.last_mut()?.children.push(Node { name: local(e.name().as_ref()), attrs, children: vec![], text: String::new(), span: (start, end) });
            }
            Event::Text(t) => {
                if let Ok(text) = t.unescape() {
                    stack.last_mut()?.text.push_str(&text);
                }
            }
            Event::End(_) => {
                let mut node = stack.pop()?;
                node.span.1 = reader.buffer_position() as usize;
                stack.last_mut()?.children.push(node);
            }
            Event::Eof => break,
            _ => {}
        }
    }
    stack.pop()?.children.into_iter().next()
}

/// The colour of a fill/line/font element: its srgbClr, schemeClr or sysClr
/// child with modifiers applied. `None` for noFill or no colour at all.
pub(super) fn color_of(parent: &Node, xml: &str, theme: &ThemeColors) -> Option<String> {
    let c = parent.children.iter().find(|c| matches!(c.name.as_str(), "srgbClr" | "schemeClr" | "sysClr" | "prstClr"))?;
    let base = match c.name.as_str() {
        "srgbClr" => c.attr("val")?.to_string(),
        "sysClr" => c.attr("lastClr").unwrap_or("000000").to_string(),
        "schemeClr" => theme.resolve(c.attr("val")?)?.clone(),
        _ => match c.attr("val")? {
            "white" => "FFFFFF".to_string(),
            _ => "000000".to_string(),
        },
    };
    if c.children.is_empty() {
        return Some(base.to_ascii_uppercase());
    }
    // Re-read the element so the shared parser applies its modifiers exactly
    // as it does for a Word drawing.
    let mut reader = Reader::from_str(&xml[c.span.0..c.span.1]);
    loop {
        match reader.read_event().ok()? {
            Event::Start(_) => break,
            Event::Eof => return Some(base),
            _ => {}
        }
    }
    Some(super::ooxml::parse_color_modifiers(&mut reader, &base, &c.name))
}

/// The mean of hex colours: a gradient drawn as one colour.
pub(super) fn average(colors: impl Iterator<Item = String>) -> Option<String> {
    let rgb: Vec<[u32; 3]> = colors
        .filter_map(|c| {
            let v = u32::from_str_radix(c.get(..6)?, 16).ok()?;
            Some([v >> 16, (v >> 8) & 0xFF, v & 0xFF])
        })
        .collect();
    if rgb.is_empty() {
        return None;
    }
    let n = rgb.len() as u32;
    let mean = |i: usize| (rgb.iter().map(|c| c[i]).sum::<u32>() + n / 2) / n;
    Some(format!("{:02X}{:02X}{:02X}", mean(0), mean(1), mean(2)))
}


/// A run of `text` in `style`, with nothing else attached.
pub(super) fn run(text: String, style: RunStyle) -> Run {
    Run {
        text,
        style,
        url: None,
        footnote_ref: None,
        endnote_ref: None,
        comment_range_start: Vec::new(),
        comment_range_end: Vec::new(),
        comment_references: Vec::new(),
        tracked_change: None,
        rpr_change: None,
        ruby: None,
        bookmark_name: None,
        is_math: false,
        field_type: None,
        has_last_rendered_page_break: false,
    }
}

/// A group member text box at `rect` (x, y, w, h in points, relative to the
/// group), with insets clockwise from left as [left, top, right, bottom],
/// a vertical anchor ("top", "middle", "bottom") and no fill or outline.
pub(super) fn text_box(
    blocks: Vec<Block>,
    insets: [f32; 4],
    v_anchor: &str,
    wrap_none: bool,
    (x, y, w, h): (f32, f32, f32, f32),
) -> TextBox {
    TextBox {
        wrap_polygon: Vec::new(),
        text_outline_inset: 0.0,
        vertical_text: false,
        blocks,
        inline: false,
        font_ref_color: None,
        auto_fit: false,
        width: w,
        height: h,
        position: Some(FloatingPosition {
            flow_page_offset: 0,
            flow_boundary_offset: 0,
            flow_source: None,
            x,
            y,
            h_relative: Some("group".to_string()),
            v_relative: Some("group".to_string()),
            h_align: None,
            v_align: None,
            dist_l: None,
            dist_r: None, dist_t: None, dist_b: None, eff_l: 0.0, eff_r: 0.0, eff_t: 0.0, eff_b: 0.0,
        }),
        border: false,
        stroke_color: None,
        stroke_width: None,
        fill: None,
        anchor_block_index: 0,
        corner_radius: None,
        ellipse: false,
        callout: None,
        outline_cmpd: None,
        outer_shadow: None,
        outline_dash: None,
        inset_left: Some(insets[0]),
        inset_right: Some(insets[2]),
        inset_top: Some(insets[1]),
        inset_bottom: Some(insets[3]),
        wrap_type: Some(WrapType::None),
        v_text_anchor: Some(v_anchor.to_string()),
        relative_height: 0,
        behind_doc: false,
        vert_overflow: None,
        compat_line_spacing: false,
        vector_shapes: Vec::new(),
        host_break_after: false,
        children: Vec::new(),
        body_wrap_none: wrap_none,
        children_on_line: false,
        chain: None,
        canvas: false,
    }
}
