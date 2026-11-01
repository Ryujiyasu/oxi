// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! SmartArt, drawn from the laid-out shapes Word stores beside it.
//!
//! A SmartArt graphic is a data model plus a layout algorithm, and running
//! that algorithm is Word's business. Word also saves its result: the
//! `diagrams/drawingN.xml` part (`dsp:drawing`) holds every shape already
//! placed, filled and lettered. That part is what this module reads. Each
//! `dsp:sp` becomes a vector shape of the group and, when it carries text, a
//! child text box over the shape's text rectangle (`dsp:txXfrm`) -- the same
//! form a Word drawing group takes, so the layout draws it the same way.

use quick_xml::events::Event;
use quick_xml::Reader;

use super::theme::ThemeColors;
use crate::ir::{
    Alignment, Block, FloatingPosition, Paragraph, ParagraphStyle, PathSeg, Run, RunStyle, TextBox,
    VectorShape, WrapType,
};

const EMU_PER_PT: f32 = 12700.0;

/// The shapes and text boxes of one SmartArt, positioned relative to its
/// frame's top-left in points.
pub(super) struct SmartArt {
    pub shapes: Vec<VectorShape>,
    pub texts: Vec<TextBox>,
}

/// One element of the drawing part, with the byte span it came from (a colour
/// element's modifiers are read again from its source by the shared colour
/// parser).
struct Node {
    name: String,
    attrs: Vec<(String, String)>,
    children: Vec<Node>,
    text: String,
    span: (usize, usize),
}

impl Node {
    fn attr(&self, key: &str) -> Option<&str> {
        self.attrs.iter().find(|(k, _)| k == key).map(|(_, v)| v.as_str())
    }
    fn child(&self, name: &str) -> Option<&Node> {
        self.children.iter().find(|c| c.name == name)
    }
    fn path(&self, names: &[&str]) -> Option<&Node> {
        names.iter().try_fold(self, |n, name| n.child(name))
    }
    fn num(&self, key: &str) -> Option<f32> {
        self.attr(key)?.parse().ok()
    }
}

fn local(name: &[u8]) -> String {
    let s = String::from_utf8_lossy(name);
    s.rsplit(':').next().unwrap_or_default().to_string()
}

fn parse_tree(xml: &str) -> Option<Node> {
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
fn color_of(parent: &Node, xml: &str, theme: &ThemeColors) -> Option<String> {
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

/// A solid fill, from the shape's own properties or, without any, its style's
/// fillRef. A gradient is drawn in its first stop's colour.
fn fill_of(sp: &Node, xml: &str, theme: &ThemeColors) -> Option<String> {
    let pr = sp.child("spPr")?;
    if pr.child("noFill").is_some() {
        return None;
    }
    if let Some(f) = pr.child("solidFill") {
        return color_of(f, xml, theme);
    }
    if let Some(stops) = pr.path(&["gradFill", "gsLst"]) {
        return average(stops.children.iter().filter_map(|gs| color_of(gs, xml, theme)));
    }
    let r = sp.path(&["style", "fillRef"])?;
    (r.num("idx").unwrap_or(0.0) > 0.0).then(|| color_of(r, xml, theme)).flatten()
}

/// The mean of hex colours: a gradient drawn as one colour.
fn average(colors: impl Iterator<Item = String>) -> Option<String> {
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

/// The outline colour and width in points.
fn line_of(sp: &Node, xml: &str, theme: &ThemeColors) -> (Option<String>, f32) {
    let ln = sp.path(&["spPr", "ln"]);
    let width = ln.and_then(|l| l.num("w")).map(|w| w / EMU_PER_PT);
    if let Some(ln) = ln {
        if ln.child("noFill").is_some() {
            return (None, 0.0);
        }
        if let Some(f) = ln.child("solidFill") {
            return (color_of(f, xml, theme), width.unwrap_or(0.75));
        }
    }
    let Some(r) = sp.path(&["style", "lnRef"]) else { return (None, 0.0) };
    let idx = r.num("idx").unwrap_or(0.0) as usize;
    if idx == 0 {
        return (None, 0.0);
    }
    let w = width.or_else(|| theme.ln_style_widths.get(idx - 1).copied()).unwrap_or(0.75);
    (color_of(r, xml, theme), w)
}

/// An adjust value of a preset (`<a:gd name="adj" fmla="val 50000"/>`), as
/// a fraction, or `default`.
fn adj(sp: &Node, name: &str, default: f32) -> f32 {
    sp.path(&["spPr", "prstGeom", "avLst"])
        .and_then(|av| av.children.iter().find(|g| g.attr("name") == Some(name)))
        .and_then(|g| g.attr("fmla")?.strip_prefix("val ")?.trim().parse::<f32>().ok())
        .map_or(default, |v| v / 100_000.0)
}

/// The outline of a preset shape in absolute points, before rotation.
fn preset_outline(sp: &Node, prst: &str, x: f32, y: f32, w: f32, h: f32) -> Vec<PathSeg> {
    let ss = w.min(h);
    let poly = |pts: &[(f32, f32)]| -> Vec<PathSeg> {
        let mut out: Vec<PathSeg> = pts.iter().enumerate()
            .map(|(i, &(px, py))| if i == 0 { PathSeg::M(x + px, y + py) } else { PathSeg::L(x + px, y + py) })
            .collect();
        out.push(PathSeg::Z);
        out
    };
    let k = 0.552_284_8;
    match prst {
        "ellipse" => {
            let (cx, cy, rx, ry) = (x + w / 2.0, y + h / 2.0, w / 2.0, h / 2.0);
            vec![
                PathSeg::M(cx + rx, cy),
                PathSeg::C(cx + rx, cy + ry * k, cx + rx * k, cy + ry, cx, cy + ry),
                PathSeg::C(cx - rx * k, cy + ry, cx - rx, cy + ry * k, cx - rx, cy),
                PathSeg::C(cx - rx, cy - ry * k, cx - rx * k, cy - ry, cx, cy - ry),
                PathSeg::C(cx + rx * k, cy - ry, cx + rx, cy - ry * k, cx + rx, cy),
                PathSeg::Z,
            ]
        }
        // Corner radii clockwise from top-left, as fractions of the short side.
        "roundRect" => rounded(x, y, w, h, [adj(sp, "adj", 0.16667); 4]),
        "round1Rect" => rounded(x, y, w, h, [0.0, adj(sp, "adj", 0.16667), 0.0, 0.0]),
        "round2SameRect" => {
            let (a, b) = (adj(sp, "adj1", 0.16667), adj(sp, "adj2", 0.0));
            rounded(x, y, w, h, [a, a, b, b])
        }
        "round2DiagRect" => {
            let (a, b) = (adj(sp, "adj1", 0.16667), adj(sp, "adj2", 0.0));
            rounded(x, y, w, h, [a, b, a, b])
        }
        "triangle" => poly(&[(w * adj(sp, "adj", 0.5), 0.0), (w, h), (0.0, h)]),
        "diamond" => poly(&[(w / 2.0, 0.0), (w, h / 2.0), (w / 2.0, h), (0.0, h / 2.0)]),
        "chevron" => {
            let d = (adj(sp, "adj", 0.5) * ss).min(w);
            poly(&[(0.0, 0.0), (w - d, 0.0), (w, h / 2.0), (w - d, h), (0.0, h), (d, h / 2.0)])
        }
        "homePlate" => {
            let d = (adj(sp, "adj", 0.5) * ss).min(w);
            poly(&[(0.0, 0.0), (w - d, 0.0), (w, h / 2.0), (w - d, h), (0.0, h)])
        }
        "pentagon" => poly(&[(w / 2.0, 0.0), (w, h * 0.382), (w * 0.809, h), (w * 0.191, h), (0.0, h * 0.382)]),
        "hexagon" => {
            let d = (adj(sp, "adj", 0.25) * ss).min(w / 2.0);
            poly(&[(d, 0.0), (w - d, 0.0), (w, h / 2.0), (w - d, h), (d, h), (0.0, h / 2.0)])
        }
        "rightArrow" | "leftArrow" => {
            let shaft = adj(sp, "adj1", 0.5) * h;
            let head = (adj(sp, "adj2", 0.5) * ss).min(w);
            let (t, b) = ((h - shaft) / 2.0, (h + shaft) / 2.0);
            let right = [(0.0, t), (w - head, t), (w - head, 0.0), (w, h / 2.0), (w - head, h), (w - head, b), (0.0, b)];
            if prst == "rightArrow" {
                poly(&right)
            } else {
                poly(&right.map(|(px, py)| (w - px, py)))
            }
        }
        "upArrow" | "downArrow" => {
            let shaft = adj(sp, "adj1", 0.5) * w;
            let head = (adj(sp, "adj2", 0.5) * ss).min(h);
            let (l, r) = ((w - shaft) / 2.0, (w + shaft) / 2.0);
            let down = [(l, 0.0), (r, 0.0), (r, h - head), (w, h - head), (w / 2.0, h), (0.0, h - head), (l, h - head)];
            if prst == "downArrow" {
                poly(&down)
            } else {
                poly(&down.map(|(px, py)| (px, h - py)))
            }
        }
        _ => poly(&[(0.0, 0.0), (w, 0.0), (w, h), (0.0, h)]),
    }
}

/// A rectangle whose corners, clockwise from top-left, are rounded by
/// `radii` times its short side.
fn rounded(x: f32, y: f32, w: f32, h: f32, radii: [f32; 4]) -> Vec<PathSeg> {
    let ss = w.min(h);
    let [tl, tr, br, bl] = radii.map(|f| (f * ss).clamp(0.0, ss / 2.0));
    let k = 1.0 - 0.552_284_8;
    let (x1, y1) = (x + w, y + h);
    let mut out = vec![PathSeg::M(x + tl, y), PathSeg::L(x1 - tr, y)];
    if tr > 0.0 { out.push(PathSeg::C(x1 - tr * k, y, x1, y + tr * k, x1, y + tr)); }
    out.push(PathSeg::L(x1, y1 - br));
    if br > 0.0 { out.push(PathSeg::C(x1, y1 - br * k, x1 - br * k, y1, x1 - br, y1)); }
    out.push(PathSeg::L(x + bl, y1));
    if bl > 0.0 { out.push(PathSeg::C(x + bl * k, y1, x, y1 - bl * k, x, y1 - bl)); }
    out.push(PathSeg::L(x, y + tl));
    if tl > 0.0 { out.push(PathSeg::C(x, y + tl * k, x + tl * k, y, x + tl, y)); }
    out.push(PathSeg::Z);
    out
}

/// A custGeom outline in absolute points: moveTo, lnTo, cubicBezTo,
/// quadBezTo, arcTo and close, each path scaled from its own w/h to the box.
fn custom_outline(geom: &Node, x: f32, y: f32, w: f32, h: f32) -> Vec<PathSeg> {
    let mut out = Vec::new();
    for path in geom.path(&["pathLst"]).map(|p| p.children.as_slice()).unwrap_or_default() {
        let pw = path.num("w").filter(|v| *v > 0.0).unwrap_or(w * EMU_PER_PT);
        let ph = path.num("h").filter(|v| *v > 0.0).unwrap_or(h * EMU_PER_PT);
        let (sx, sy) = (w / pw, h / ph);
        let pt = |n: &Node| -> Option<(f32, f32)> { Some((x + n.num("x")? * sx, y + n.num("y")? * sy)) };
        let mut cur = (x, y);
        for cmd in &path.children {
            let pts: Vec<(f32, f32)> = cmd.children.iter().filter_map(&pt).collect();
            match (cmd.name.as_str(), pts.as_slice()) {
                ("moveTo", [p]) => { out.push(PathSeg::M(p.0, p.1)); cur = *p; }
                ("lnTo", [p]) => { out.push(PathSeg::L(p.0, p.1)); cur = *p; }
                ("cubicBezTo", [a, b, c]) => { out.push(PathSeg::C(a.0, a.1, b.0, b.1, c.0, c.1)); cur = *c; }
                ("quadBezTo", [a, c]) => {
                    let c1 = (cur.0 + 2.0 / 3.0 * (a.0 - cur.0), cur.1 + 2.0 / 3.0 * (a.1 - cur.1));
                    let c2 = (c.0 + 2.0 / 3.0 * (a.0 - c.0), c.1 + 2.0 / 3.0 * (a.1 - c.1));
                    out.push(PathSeg::C(c1.0, c1.1, c2.0, c2.1, c.0, c.1));
                    cur = *c;
                }
                ("arcTo", _) => {
                    let (Some(wr), Some(hr), Some(st), Some(sw)) = (cmd.num("wR"), cmd.num("hR"), cmd.num("stAng"), cmd.num("swAng")) else { continue };
                    cur = arc_to(&mut out, cur, wr * sx, hr * sy, st / 60_000.0, sw / 60_000.0);
                }
                ("close", _) => out.push(PathSeg::Z),
                _ => {}
            }
        }
    }
    out
}

/// DrawingML arcTo: an elliptical arc of radii (rx, ry) starting at `cur`,
/// which lies on the ellipse at `start` degrees, sweeping `sweep` degrees.
/// Emitted as cubic segments of at most 90°. Returns the end point.
fn arc_to(out: &mut Vec<PathSeg>, cur: (f32, f32), rx: f32, ry: f32, start: f32, sweep: f32) -> (f32, f32) {
    let on = |deg: f32| {
        // DrawingML angles are of the visual point, not the parametric one.
        let a = deg.to_radians();
        let t = (a.sin() * rx).atan2(a.cos() * ry);
        (rx * t.cos(), ry * t.sin(), t)
    };
    let (sx, sy, t0) = on(start);
    let (cx, cy) = (cur.0 - sx, cur.1 - sy);
    let (_, _, t_end) = on(start + sweep);
    let mut dt = t_end - t0;
    if sweep > 0.0 && dt <= 0.0 { dt += std::f32::consts::TAU; }
    if sweep < 0.0 && dt >= 0.0 { dt -= std::f32::consts::TAU; }
    if sweep.abs() >= 360.0 { dt = sweep.signum() * std::f32::consts::TAU; }
    let n = (dt.abs() / std::f32::consts::FRAC_PI_2).ceil().max(1.0) as usize;
    let step = dt / n as f32;
    let k = 4.0 / 3.0 * (step / 4.0).tan();
    let mut t = t0;
    let mut end = cur;
    for _ in 0..n {
        let t2 = t + step;
        let (p0x, p0y) = (cx + rx * t.cos(), cy + ry * t.sin());
        let (p3x, p3y) = (cx + rx * t2.cos(), cy + ry * t2.sin());
        out.push(PathSeg::C(
            p0x - k * rx * t.sin(), p0y + k * ry * t.cos(),
            p3x + k * rx * t2.sin(), p3y - k * ry * t2.cos(),
            p3x, p3y,
        ));
        end = (p3x, p3y);
        t = t2;
    }
    end
}

/// Apply the shape's flips and rotation about its centre, then express the
/// outline over its box as the 0..1 coordinates `VectorShape::path` uses.
fn place(segs: Vec<PathSeg>, (x, y, w, h): (f32, f32, f32, f32), rot: f32, flip_h: bool, flip_v: bool) -> Vec<PathSeg> {
    let (cx, cy) = (x + w / 2.0, y + h / 2.0);
    let (s, c) = rot.to_radians().sin_cos();
    let tf = |px: f32, py: f32| {
        let px = if flip_h { 2.0 * cx - px } else { px };
        let py = if flip_v { 2.0 * cy - py } else { py };
        let (dx, dy) = (px - cx, py - cy);
        let (rx, ry) = (cx + dx * c - dy * s, cy + dx * s + dy * c);
        ((rx - x) / w.max(f32::EPSILON), (ry - y) / h.max(f32::EPSILON))
    };
    segs.into_iter()
        .map(|seg| match seg {
            PathSeg::M(a, b) => { let p = tf(a, b); PathSeg::M(p.0, p.1) }
            PathSeg::L(a, b) => { let p = tf(a, b); PathSeg::L(p.0, p.1) }
            PathSeg::C(a, b, c2, d, e, f) => {
                let (p, q, r) = (tf(a, b), tf(c2, d), tf(e, f));
                PathSeg::C(p.0, p.1, q.0, q.1, r.0, r.1)
            }
            PathSeg::Z => PathSeg::Z,
        })
        .collect()
}

fn xfrm(n: &Node) -> Option<(f32, f32, f32, f32)> {
    let off = n.child("off")?;
    let ext = n.child("ext")?;
    Some((
        off.num("x")? / EMU_PER_PT,
        off.num("y")? / EMU_PER_PT,
        ext.num("cx")? / EMU_PER_PT,
        ext.num("cy")? / EMU_PER_PT,
    ))
}

/// The text of one shape as paragraphs, and its body's insets and anchor.
fn text_blocks(sp: &Node, xml: &str, theme: &ThemeColors) -> Vec<Block> {
    let Some(body) = sp.child("txBody") else { return Vec::new() };
    let default_color = sp.path(&["style", "fontRef"]).and_then(|f| color_of(f, xml, theme));
    let mut blocks = Vec::new();
    for p in body.children.iter().filter(|c| c.name == "p") {
        let ppr = p.child("pPr");
        let alignment = match ppr.and_then(|pr| pr.attr("algn")) {
            Some("ctr") => Alignment::Center,
            Some("r") => Alignment::Right,
            Some("just") => Alignment::Justify,
            Some("dist") => Alignment::Distribute,
            _ => Alignment::Left,
        };
        let mut runs = Vec::new();
        let mut size_pt: f32 = 18.0;
        for r in &p.children {
            let text = match r.name.as_str() {
                "r" | "fld" => r.child("t").map(|t| t.text.clone()).unwrap_or_default(),
                "br" => "\n".to_string(),
                _ => continue,
            };
            let rpr = r.child("rPr");
            if let Some(sz) = rpr.and_then(|pr| pr.num("sz")) {
                size_pt = sz / 100.0;
            }
            let face = |slot: &str| rpr.and_then(|pr| pr.child(slot)).and_then(|f| f.attr("typeface")).map(str::to_string);
            let theme_face = |f: Option<String>, minor: &Option<String>| match f.as_deref() {
                None | Some("+mn-lt") | Some("+mn-ea") => minor.clone(),
                Some("+mj-lt") => theme.major_font.clone(),
                Some("+mj-ea") => theme.major_font_ea.clone(),
                _ => f,
            };
            let color = rpr.and_then(|pr| pr.child("solidFill")).and_then(|f| color_of(f, xml, theme)).or_else(|| default_color.clone());
            runs.push(Run {
                text,
                style: RunStyle {
                    font_size: Some(size_pt),
                    bold: rpr.and_then(|pr| pr.attr("b")) == Some("1"),
                    italic: rpr.and_then(|pr| pr.attr("i")) == Some("1"),
                    underline: rpr.and_then(|pr| pr.attr("u")).is_some_and(|u| u != "none"),
                    color,
                    font_family: theme_face(face("latin"), &theme.minor_font),
                    font_family_east_asia: theme_face(face("ea"), &theme.minor_font_ea),
                    ..RunStyle::default()
                },
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
            });
        }
        if runs.is_empty() {
            size_pt = p.child("endParaRPr").and_then(|e| e.num("sz")).map_or(size_pt, |s| s / 100.0);
        }
        let spacing = |name: &str| -> Option<f32> {
            let s = ppr?.child(name)?;
            if let Some(pct) = s.child("spcPct").and_then(|v| v.num("val")) {
                Some(pct / 100_000.0 * size_pt)
            } else {
                s.child("spcPts").and_then(|v| v.num("val")).map(|v| v / 100.0)
            }
        };
        let line = ppr.and_then(|pr| pr.path(&["lnSpc", "spcPct"])).and_then(|v| v.num("val")).map(|v| v / 100_000.0);
        blocks.push(Block::Paragraph(Paragraph {
            runs,
            style: ParagraphStyle {
                space_before: Some(spacing("spcBef").unwrap_or(0.0)),
                space_after: Some(spacing("spcAft").unwrap_or(0.0)),
                line_spacing: line,
                snap_to_grid: false,
                ..ParagraphStyle::default()
            },
            alignment,
            shapes: Vec::new(),
            ppr_change: None,
            paragraph_mark_revision: None,
        }));
    }
    blocks
}

fn text_box(blocks: Vec<Block>, body: Option<&Node>, (x, y, w, h): (f32, f32, f32, f32)) -> TextBox {
    let inset = |key: &str, default: f32| body.and_then(|b| b.num(key)).map_or(default, |v| v / EMU_PER_PT);
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
        inset_left: Some(inset("lIns", 7.2)),
        inset_right: Some(inset("rIns", 7.2)),
        inset_top: Some(inset("tIns", 3.6)),
        inset_bottom: Some(inset("bIns", 3.6)),
        wrap_type: Some(WrapType::None),
        v_text_anchor: match body.and_then(|b| b.attr("anchor")) {
            Some("ctr") => Some("middle".to_string()),
            Some("b") => Some("bottom".to_string()),
            _ => Some("top".to_string()),
        },
        relative_height: 0,
        behind_doc: false,
        vert_overflow: None,
        compat_line_spacing: false,
        vector_shapes: Vec::new(),
        host_break_after: false,
        children: Vec::new(),
        body_wrap_none: body.and_then(|b| b.attr("wrap")) == Some("none"),
        children_on_line: false,
    }
}

/// The SmartArt drawn by `xml`, a `dsp:drawing` part.
pub(super) fn from_drawing(xml: &str, theme: &ThemeColors) -> Option<SmartArt> {
    let root = parse_tree(xml)?;
    let tree = root.child("spTree")?;
    let mut art = SmartArt { shapes: Vec::new(), texts: Vec::new() };
    let mut stack: Vec<&Node> = tree.children.iter().rev().collect();
    while let Some(n) = stack.pop() {
        match n.name.as_str() {
            "grpSp" => stack.extend(n.children.iter().rev()),
            "sp" => {
                let Some(x) = n.path(&["spPr", "xfrm"]) else { continue };
                let Some(rect) = xfrm(x) else { continue };
                let rot = x.num("rot").unwrap_or(0.0) / 60_000.0;
                let (flip_h, flip_v) = (x.attr("flipH") == Some("1"), x.attr("flipV") == Some("1"));
                let fill = fill_of(n, xml, theme);
                let (stroke, stroke_width) = line_of(n, xml, theme);
                if fill.is_some() || stroke.is_some() {
                    let (bx, by, bw, bh) = rect;
                    let outline = if let Some(g) = n.path(&["spPr", "custGeom"]) {
                        custom_outline(g, bx, by, bw, bh)
                    } else {
                        let prst = n.path(&["spPr", "prstGeom"]).and_then(|g| g.attr("prst")).unwrap_or("rect");
                        preset_outline(n, prst, bx, by, bw, bh)
                    };
                    if !outline.is_empty() {
                        art.shapes.push(VectorShape {
                            x: bx,
                            y: by,
                            w: bw,
                            h: bh,
                            fill,
                            stroke,
                            stroke_width,
                            is_line: false,
                            path: place(outline, rect, rot, flip_h, flip_v),
                        });
                    }
                }
                let blocks = text_blocks(n, xml, theme);
                let has_text = blocks.iter().any(|b| matches!(b, Block::Paragraph(p) if p.runs.iter().any(|r| !r.text.trim().is_empty())));
                if has_text {
                    let area = n.child("txXfrm").and_then(xfrm).unwrap_or(rect);
                    art.texts.push(text_box(blocks, n.path(&["txBody", "bodyPr"]), area));
                }
            }
            _ => {}
        }
    }
    (!art.shapes.is_empty() || !art.texts.is_empty()).then_some(art)
}

#[cfg(test)]
mod tests {
    use super::*;

    const DRAWING: &str = r#"<dsp:drawing xmlns:dsp="d" xmlns:a="a"><dsp:spTree><dsp:nvGrpSpPr/><dsp:grpSpPr/>
<dsp:sp><dsp:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1270000" cy="635000"/></a:xfrm><a:prstGeom prst="roundRect"><a:avLst/></a:prstGeom><a:solidFill><a:schemeClr val="accent1"/></a:solidFill><a:ln w="25400"><a:solidFill><a:srgbClr val="ffffff"/></a:solidFill></a:ln></dsp:spPr>
<dsp:style><a:lnRef idx="2"><a:scrgbClr r="0" g="0" b="0"/></a:lnRef><a:fillRef idx="1"><a:scrgbClr r="0" g="0" b="0"/></a:fillRef><a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef></dsp:style>
<dsp:txBody><a:bodyPr lIns="12700" tIns="12700" rIns="12700" bIns="12700" anchor="ctr"/><a:p><a:pPr algn="ctr"/><a:r><a:rPr lang="ja-JP" sz="2300" b="1"/><a:t>計画</a:t></a:r></a:p></dsp:txBody>
<dsp:txXfrm><a:off x="12700" y="12700"/><a:ext cx="1244600" cy="609600"/></dsp:txXfrm></dsp:sp>
<dsp:sp><dsp:spPr><a:xfrm rot="5400000"><a:off x="1270000" y="0"/><a:ext cx="635000" cy="317500"/></a:xfrm><a:prstGeom prst="rightArrow"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="808080"/></a:solidFill><a:ln><a:noFill/></a:ln></dsp:spPr></dsp:sp>
</dsp:spTree></dsp:drawing>"#;

    fn theme() -> ThemeColors {
        let mut t = ThemeColors::default();
        t.colors.insert("accent1".into(), "4472C4".into());
        t.colors.insert("lt1".into(), "FFFFFF".into());
        t.minor_font = Some("Calibri".into());
        t.minor_font_ea = Some("游明朝".into());
        t
    }

    #[test]
    fn shapes_and_text_come_from_the_drawing() {
        let art = from_drawing(DRAWING, &theme()).expect("smartart");
        assert_eq!(art.shapes.len(), 2);
        let box1 = &art.shapes[0];
        assert_eq!((box1.x, box1.y, box1.w, box1.h), (0.0, 0.0, 100.0, 50.0));
        assert_eq!(box1.fill.as_deref(), Some("4472C4"));
        assert_eq!(box1.stroke.as_deref(), Some("FFFFFF"));
        assert_eq!(box1.stroke_width, 2.0);
        // a rounded rectangle is an outline with curved corners, inside its box
        assert!(box1.path.iter().any(|s| matches!(s, PathSeg::C(..))));
        // no outline on the arrow
        assert_eq!(art.shapes[1].stroke, None);

        assert_eq!(art.texts.len(), 1, "only the shape with text gets a text box");
        let t = &art.texts[0];
        let pos = t.position.as_ref().unwrap();
        assert_eq!((pos.x, pos.y, t.width, t.height), (1.0, 1.0, 98.0, 48.0));
        assert_eq!(t.v_text_anchor.as_deref(), Some("middle"));
        assert_eq!(t.inset_left, Some(1.0));
        let Block::Paragraph(p) = &t.blocks[0] else { panic!() };
        assert!(matches!(p.alignment, Alignment::Center));
        let r = &p.runs[0];
        assert_eq!(r.text, "計画");
        assert_eq!(r.style.font_size, Some(23.0));
        assert!(r.style.bold);
        assert_eq!(r.style.color.as_deref(), Some("FFFFFF"), "text takes the style's fontRef colour");
        assert_eq!(r.style.font_family_east_asia.as_deref(), Some("游明朝"));
    }

    #[test]
    fn rotation_turns_the_outline_about_the_centre() {
        let art = from_drawing(DRAWING, &theme()).unwrap();
        // the arrow is 50x25 turned a quarter: its tip, at the right middle
        // before turning, ends at the bottom middle
        let tip = art.shapes[1].path.iter().filter_map(|s| match s { PathSeg::L(x, y) => Some((*x, *y)), _ => None })
            .find(|(x, _)| (x - 0.5).abs() < 1e-3).expect("a point on the vertical centre line");
        assert!((tip.1 - 1.5).abs() < 1e-3, "tip at {tip:?}");
    }

    /// Strict OOXML writes modifier percentages as "60%". Read as 0, the tint
    /// washed the colour out to white and the shape vanished.
    #[test]
    fn strict_percentages_are_read() {
        let strict = DRAWING.replace(
            r#"<a:schemeClr val="accent1"/></a:solidFill><a:ln w="25400">"#,
            r#"<a:schemeClr val="accent1"><a:tint val="60%"/></a:schemeClr></a:solidFill><a:ln w="25400">"#,
        );
        let transitional = DRAWING.replace(
            r#"<a:schemeClr val="accent1"/></a:solidFill><a:ln w="25400">"#,
            r#"<a:schemeClr val="accent1"><a:tint val="60000"/></a:schemeClr></a:solidFill><a:ln w="25400">"#,
        );
        let a = from_drawing(&strict, &theme()).unwrap().shapes[0].fill.clone();
        let b = from_drawing(&transitional, &theme()).unwrap().shapes[0].fill.clone();
        assert_eq!(a, b);
        assert_ne!(a.as_deref(), Some("FFFFFF"));
    }

    #[test]
    fn arc_to_draws_a_quarter_circle() {
        let mut out = Vec::new();
        let end = arc_to(&mut out, (10.0, 0.0), 10.0, 10.0, 0.0, 90.0);
        assert!((end.0 - 0.0).abs() < 1e-3 && (end.1 - 10.0).abs() < 1e-3, "{end:?}");
        assert_eq!(out.len(), 1);
    }
}
