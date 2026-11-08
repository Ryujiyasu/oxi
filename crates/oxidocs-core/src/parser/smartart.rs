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

use super::dml::{average, color_of, parse_tree, run, text_box, Group, Node, EMU_PER_PT};
use super::theme::ThemeColors;
use crate::ir::{Alignment, Block, Paragraph, ParagraphStyle, PathSeg, RunStyle, TextBox, VectorShape};

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
    text_blocks_in(sp.child("txBody"), sp, xml, theme)
}

/// The paragraphs of the text body `body` of shape `sp`.
fn text_blocks_in(body: Option<&Node>, sp: &Node, xml: &str, theme: &ThemeColors) -> Vec<Block> {
    let Some(body) = body else { return Vec::new() };
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
            runs.push(run(
                text,
                RunStyle {
                font_size: Some(size_pt),
                bold: rpr.and_then(|pr| pr.attr("b")) == Some("1"),
                italic: rpr.and_then(|pr| pr.attr("i")) == Some("1"),
                underline: rpr.and_then(|pr| pr.attr("u")).is_some_and(|u| u != "none"),
                color,
                font_family: theme_face(face("latin"), &theme.minor_font),
                font_family_east_asia: theme_face(face("ea"), &theme.minor_font_ea),
                ..RunStyle::default()
            },
            ));
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

fn body_text_box(blocks: Vec<Block>, body: Option<&Node>, rect: (f32, f32, f32, f32)) -> TextBox {
    let inset = |key: &str, default: f32| body.and_then(|b| b.num(key)).map_or(default, |v| v / EMU_PER_PT);
    let anchor = match body.and_then(|b| b.attr("anchor")) {
        Some("ctr") => "middle",
        Some("b") => "bottom",
        _ => "top",
    };
    text_box(
        blocks,
        [inset("lIns", 7.2), inset("tIns", 3.6), inset("rIns", 7.2), inset("bIns", 3.6)],
        anchor,
        body.and_then(|b| b.attr("wrap")) == Some("none"),
        rect,
    )
}

/// The SmartArt drawn by `xml`, a `dsp:drawing` part.
pub(super) fn from_drawing(xml: &str, theme: &ThemeColors) -> Option<Group> {
    let root = parse_tree(xml)?;
    let tree = root.child("spTree")?;
    let mut art = Group { shapes: Vec::new(), texts: Vec::new() };
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
                            arrow: None, image: None,
                        });
                    }
                }
                let blocks = text_blocks(n, xml, theme);
                let has_text = blocks.iter().any(|b| matches!(b, Block::Paragraph(p) if p.runs.iter().any(|r| !r.text.trim().is_empty())));
                if has_text {
                    let area = n.child("txXfrm").and_then(xfrm).unwrap_or(rect);
                    art.texts.push(body_text_box(blocks, n.path(&["txBody", "bodyPr"]), area));
                }
            }
            _ => {}
        }
    }
    (!art.shapes.is_empty() || !art.texts.is_empty()).then_some(art)
}

/// One group's child-to-parent transform, in points: its children's
/// `chOff`/`chExt` box is drawn over the group's `off`/`ext` box.
#[derive(Clone, Copy, Debug)]
struct GroupXf {
    off: (f32, f32),
    ext: (f32, f32),
    ch_off: (f32, f32),
    ch_ext: (f32, f32),
}

impl GroupXf {
    fn of(grp_sp_pr: &Node) -> Option<Self> {
        let x = grp_sp_pr.child("xfrm")?;
        let pair = |name: &str, a: &str, b: &str| -> Option<(f32, f32)> {
            let n = x.child(name)?;
            Some((n.num(a)? / EMU_PER_PT, n.num(b)? / EMU_PER_PT))
        };
        let ext = pair("ext", "cx", "cy")?;
        Some(Self {
            off: pair("off", "x", "y").unwrap_or((0.0, 0.0)),
            ext,
            ch_off: pair("chOff", "x", "y").unwrap_or((0.0, 0.0)),
            ch_ext: pair("chExt", "cx", "cy").unwrap_or(ext),
        })
    }
    fn scale(&self) -> (f32, f32) {
        (self.ext.0 / self.ch_ext.0.max(f32::EPSILON), self.ext.1 / self.ch_ext.1.max(f32::EPSILON))
    }
    fn apply(&self, (x, y, w, h): (f32, f32, f32, f32)) -> (f32, f32, f32, f32) {
        let (sx, sy) = self.scale();
        (self.off.0 + (x - self.ch_off.0) * sx, self.off.1 + (y - self.ch_off.1) * sy, w * sx, h * sy)
    }
}

/// A member's spPr xfrm and its box in its group's space.
fn member_xfrm(n: &Node) -> Option<(&Node, (f32, f32, f32, f32))> {
    let x = n.path(&["spPr", "xfrm"])?;
    Some((x, xfrm(x)?))
}

/// `rect` in a member's own group space, carried out to the frame through
/// `chain` (outermost first).
fn to_frame(chain: &[GroupXf], rect: (f32, f32, f32, f32)) -> (f32, f32, f32, f32) {
    chain.iter().rev().fold(rect, |r, xf| xf.apply(r))
}

/// The open outline of a connector preset over (x, y, w, h); `None` for a
/// straight one, which is drawn as a line.
fn connector_outline(sp: &Node, prst: &str, x: f32, y: f32, w: f32, h: f32) -> Option<Vec<PathSeg>> {
    let pts: Vec<(f32, f32)> = match prst {
        "bentConnector2" => vec![(0.0, 0.0), (w, 0.0), (w, h)],
        "bentConnector3" | "curvedConnector3" => {
            let m = adj(sp, "adj1", 0.5) * w;
            vec![(0.0, 0.0), (m, 0.0), (m, h), (w, h)]
        }
        "bentConnector4" | "curvedConnector4" => {
            let (m1, m2) = (adj(sp, "adj1", 0.5) * w, adj(sp, "adj2", 0.5) * h);
            vec![(0.0, 0.0), (m1, 0.0), (m1, m2), (w, m2), (w, h)]
        }
        _ => return None,
    };
    Some(pts.iter().enumerate()
        .map(|(i, &(px, py))| if i == 0 { PathSeg::M(x + px, y + py) } else { PathSeg::L(x + px, y + py) })
        .collect())
}

/// A locked canvas (`lc:lockedCanvas`, a group of DrawingML shapes pasted
/// from PowerPoint) drawn into a frame of `frame_w` x `frame_h` points: its
/// shapes, connectors and pictures as group members, its shape text as
/// member text boxes. The canvas is scaled to the frame as a whole, like a
/// picture: member boxes through each group's chOff/chExt -> off/ext, and
/// the text by the smaller of the two scales (UNMEASURED against Word).
/// `media` maps the document's relationship ids to image bytes.
pub(super) fn from_locked_canvas(
    xml: &str,
    theme: &ThemeColors,
    media: &std::collections::HashMap<String, Vec<u8>>,
    frame_w: f32,
    frame_h: f32,
) -> Option<Group> {
    let root = parse_tree(xml)?;
    let canvas = GroupXf::of(root.child("grpSpPr")?)?;
    if frame_w <= 0.0 || frame_h <= 0.0 {
        return None;
    }
    let frame = GroupXf { off: (0.0, 0.0), ext: (frame_w, frame_h), ch_off: canvas.off, ch_ext: canvas.ext };
    let mut art = Group { shapes: Vec::new(), texts: Vec::new() };
    // (member, the group transforms above it, outermost first)
    let mut stack: Vec<(&Node, Vec<GroupXf>)> = root.children.iter().rev().map(|n| (n, vec![frame, canvas])).collect();
    while let Some((n, chain)) = stack.pop() {
        match n.name.as_str() {
            "grpSp" => {
                let mut inner = chain.clone();
                if let Some(xf) = n.child("grpSpPr").and_then(GroupXf::of) {
                    inner.push(xf);
                }
                stack.extend(n.children.iter().rev().map(|c| (c, inner.clone())));
            }
            "pic" => {
                let Some((_, local)) = member_xfrm(n) else { continue };
                let (x, y, w, h) = to_frame(&chain, local);
                let rid = n.path(&["blipFill", "blip"]).and_then(|b| b.attr("embed"));
                if let Some(data) = rid.and_then(|id| media.get(id)).filter(|_| w > 0.0 && h > 0.0) {
                    art.shapes.push(VectorShape {
                        x, y, w, h, fill: None, stroke: None, stroke_width: 0.0, is_line: false,
                        path: Vec::new(), arrow: None, image: Some(std::sync::Arc::new(data.clone())),
                    });
                }
            }
            "cxnSp" => {
                let Some((x_node, local)) = member_xfrm(n) else { continue };
                let rect = to_frame(&chain, local);
                let (stroke, stroke_width) = line_of(n, xml, theme);
                if stroke.is_none() {
                    continue;
                }
                let rot = x_node.num("rot").unwrap_or(0.0) / 60_000.0;
                let (flip_h, flip_v) = (x_node.attr("flipH") == Some("1"), x_node.attr("flipV") == Some("1"));
                let (bx, by, bw, bh) = rect;
                let prst = n.path(&["spPr", "prstGeom"]).and_then(|g| g.attr("prst")).unwrap_or("line");
                match connector_outline(n, prst, bx, by, bw, bh) {
                    Some(open) if bw > 0.0 && bh > 0.0 => art.shapes.push(VectorShape {
                        x: bx, y: by, w: bw, h: bh, fill: None, stroke, stroke_width, is_line: false,
                        path: place(open, rect, rot, flip_h, flip_v), arrow: None, image: None,
                    }),
                    _ => {
                        // A straight connector runs corner to corner; a flip
                        // swaps the corners it joins.
                        let (x0, x1) = if flip_h { (bx + bw, bx) } else { (bx, bx + bw) };
                        let (y0, y1) = if flip_v { (by + bh, by) } else { (by, by + bh) };
                        art.shapes.push(VectorShape {
                            x: x0, y: y0, w: x1 - x0, h: y1 - y0, fill: None, stroke, stroke_width,
                            is_line: true, path: Vec::new(), arrow: None, image: None,
                        });
                    }
                }
            }
            "sp" => {
                let Some((x_node, local)) = member_xfrm(n) else { continue };
                let rect = to_frame(&chain, local);
                let rot = x_node.num("rot").unwrap_or(0.0) / 60_000.0;
                let (flip_h, flip_v) = (x_node.attr("flipH") == Some("1"), x_node.attr("flipV") == Some("1"));
                let fill = fill_of(n, xml, theme);
                let (stroke, stroke_width) = line_of(n, xml, theme);
                let (bx, by, bw, bh) = rect;
                if (fill.is_some() || stroke.is_some()) && bw > 0.0 && bh > 0.0 {
                    let outline = if let Some(g) = n.path(&["spPr", "custGeom"]) {
                        custom_outline(g, bx, by, bw, bh)
                    } else {
                        let prst = n.path(&["spPr", "prstGeom"]).and_then(|g| g.attr("prst")).unwrap_or("rect");
                        preset_outline(n, prst, bx, by, bw, bh)
                    };
                    if !outline.is_empty() {
                        art.shapes.push(VectorShape {
                            x: bx, y: by, w: bw, h: bh, fill, stroke, stroke_width, is_line: false,
                            path: place(outline, rect, rot, flip_h, flip_v), arrow: None, image: None,
                        });
                    }
                }
                // Shape text sits in a:txSp/a:txBody here (dsp:txBody in SmartArt).
                let body = n.path(&["txSp", "txBody"]).or_else(|| n.child("txBody"));
                let mut blocks = text_blocks_in(body, n, xml, theme);
                let has_text = blocks.iter().any(|b| matches!(b, Block::Paragraph(p) if p.runs.iter().any(|r| !r.text.trim().is_empty())));
                if has_text {
                    let (sx, sy) = chain.iter().fold((1.0f32, 1.0f32), |(a, b), xf| { let (u, v) = xf.scale(); (a * u, b * v) });
                    let k = sx.min(sy);
                    for b in &mut blocks {
                        if let Block::Paragraph(p) = b {
                            for r in &mut p.runs {
                                r.style.font_size = r.style.font_size.map(|s| s * k);
                            }
                            p.style.space_before = p.style.space_before.map(|s| s * k);
                            p.style.space_after = p.style.space_after.map(|s| s * k);
                        }
                    }
                    let mut tb = body_text_box(blocks, body.and_then(|b| b.child("bodyPr")), rect);
                    for inset in [&mut tb.inset_left, &mut tb.inset_top, &mut tb.inset_right, &mut tb.inset_bottom] {
                        *inset = inset.map(|v| v * k);
                    }
                    art.texts.push(tb);
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

    // A canvas of 200 x 100pt whose children start at (100, 50)pt, drawn into
    // a 100 x 50pt frame: everything halves, and the text with it.
    const CANVAS: &str = r#"<lockedCanvas><a:nvGrpSpPr/><a:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="2540000" cy="1270000"/><a:chOff x="1270000" y="635000"/><a:chExt cx="2540000" cy="1270000"/></a:xfrm></a:grpSpPr>
<a:sp><a:spPr><a:xfrm><a:off x="1270000" y="635000"/><a:ext cx="1270000" cy="635000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:spPr>
<a:txSp><a:txBody><a:bodyPr lIns="91440" anchor="ctr"/><a:p><a:r><a:rPr sz="2000"/><a:t>Team</a:t></a:r></a:p></a:txBody></a:txSp></a:sp>
<a:cxnSp><a:spPr><a:xfrm flipH="1"><a:off x="2540000" y="1270000"/><a:ext cx="1270000" cy="0"/></a:xfrm><a:prstGeom prst="line"><a:avLst/></a:prstGeom><a:ln w="25400"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln></a:spPr></a:cxnSp>
<a:grpSp><a:grpSpPr><a:xfrm><a:off x="2540000" y="635000"/><a:ext cx="1270000" cy="635000"/><a:chOff x="0" y="0"/><a:chExt cx="2540000" cy="1270000"/></a:xfrm></a:grpSpPr>
<a:pic><a:blipFill><a:blip r:embed="rId9"/></a:blipFill><a:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="2540000" cy="1270000"/></a:xfrm></a:spPr></a:pic></a:grpSp>
</lockedCanvas>"#;

    #[test]
    fn locked_canvas_members_are_scaled_into_the_frame() {
        let media = std::collections::HashMap::from([("rId9".to_string(), vec![1u8, 2, 3])]);
        let art = from_locked_canvas(CANVAS, &theme(), &media, 100.0, 50.0).expect("canvas");
        assert_eq!(art.shapes.len(), 3);
        let rect = &art.shapes[0];
        assert_eq!((rect.x, rect.y, rect.w, rect.h), (0.0, 0.0, 50.0, 25.0));
        assert_eq!(rect.fill.as_deref(), Some("FF0000"));
        // a flipped straight connector runs from its right end to its left
        let line = &art.shapes[1];
        assert!(line.is_line);
        assert_eq!((line.x, line.y, line.w, line.h), (100.0, 25.0, -50.0, 0.0));
        assert_eq!(line.stroke_width, 2.0);
        // the nested group maps its 200 x 100 children onto 100 x 50, then the frame halves it
        let pic = &art.shapes[2];
        assert_eq!((pic.x, pic.y, pic.w, pic.h), (50.0, 0.0, 50.0, 25.0));
        assert_eq!(pic.image.as_deref().map(Vec::as_slice), Some(&[1u8, 2, 3][..]));

        assert_eq!(art.texts.len(), 1);
        let t = &art.texts[0];
        assert_eq!((t.width, t.height), (50.0, 25.0));
        assert_eq!(t.inset_left, Some(3.6));
        let Block::Paragraph(p) = &t.blocks[0] else { panic!() };
        assert_eq!(p.runs[0].text, "Team");
        assert_eq!(p.runs[0].style.font_size, Some(10.0));
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
