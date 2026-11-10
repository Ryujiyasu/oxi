// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! Enhanced Metafile (EMF) pictures drawn as PDF vectors.
//!
//! OXI_PDF_EMF (2026-11-11, default ON, opt-out _DISABLE): the `image` crate
//! cannot read a metafile, so every EMF picture was left out of the PDF (en 76,
//! ja 52, zh 11 corpus documents carry one; zh dev creative__0f11c107's header
//! logos). This plays the GDI drawing records -- paths, polygons, lines,
//! rectangles, ellipses, pens, brushes, the polygon fill mode, world and
//! window/viewport transforms, saved device contexts and device-independent
//! bitmaps -- into the PDF's own path and image elements, in the picture's box.
//! Text (EXTTEXTOUTW) is set character by character at its dx advances in the
//! record's logical font; EMF+ records are not drawn.

use oxipdf_core::ir::*;

const STOCK: u32 = 0x8000_0000;

#[derive(Clone, Copy)]
struct Xform {
    m11: f64,
    m12: f64,
    m21: f64,
    m22: f64,
    dx: f64,
    dy: f64,
}

impl Xform {
    const ID: Xform = Xform { m11: 1.0, m12: 0.0, m21: 0.0, m22: 1.0, dx: 0.0, dy: 0.0 };
    fn apply(&self, x: f64, y: f64) -> (f64, f64) {
        (x * self.m11 + y * self.m21 + self.dx, x * self.m12 + y * self.m22 + self.dy)
    }
    /// self then other
    fn then(&self, o: &Xform) -> Xform {
        Xform {
            m11: self.m11 * o.m11 + self.m12 * o.m21,
            m12: self.m11 * o.m12 + self.m12 * o.m22,
            m21: self.m21 * o.m11 + self.m22 * o.m21,
            m22: self.m21 * o.m12 + self.m22 * o.m22,
            dx: self.dx * o.m11 + self.dy * o.m21 + o.dx,
            dy: self.dx * o.m12 + self.dy * o.m22 + o.dy,
        }
    }
}

#[derive(Clone)]
enum Obj {
    Pen { color: Option<Color>, width: f64 },
    Brush(Option<Color>),
    Font(LogFont),
}

#[derive(Clone)]
struct LogFont {
    height: i32,
    /// lfEscapement (tenths of a degree)
    escapement: i32,
    face: String,
    bold: bool,
    italic: bool,
}

fn logfont(rec: &[u8]) -> Option<LogFont> {
    let height = rd_i32(rec, 12)?;
    let weight = rd_i32(rec, 28)?;
    let italic = *rec.get(32)? != 0;
    let name: Vec<u16> = (0..32)
        .filter_map(|i| rec.get(40 + i * 2..42 + i * 2))
        .map(|b| u16::from_le_bytes([b[0], b[1]]))
        .take_while(|&c| c != 0)
        .collect();
    let escapement = rd_i32(rec, 20)?;
    Some(LogFont { height, escapement, face: String::from_utf16_lossy(&name), bold: weight >= 600, italic })
}

/// Every (face, bold, italic, text) the metafile sets, for the font subsetter.
pub fn text_runs(data: &[u8]) -> Vec<(String, bool, bool, String)> {
    let mut out = Vec::new();
    if !is_emf(data) {
        return out;
    }
    let mut fonts: std::collections::HashMap<u32, LogFont> = std::collections::HashMap::new();
    let mut cur: Option<LogFont> = None;
    let mut off = 0usize;
    while off + 8 <= data.len() {
        let (Some(kind), Some(size)) = (rd_u32(data, off), rd_u32(data, off + 4)) else { break };
        let size = size as usize;
        if size < 8 || off + size > data.len() || kind == 14 {
            break;
        }
        let rec = &data[off..off + size];
        match kind {
            82 => {
                if let (Some(ih), Some(f)) = (rd_u32(rec, 8), logfont(rec)) {
                    fonts.insert(ih, f);
                }
            }
            37 => {
                if let Some(f) = rd_u32(rec, 8).and_then(|ih| fonts.get(&ih)) {
                    cur = Some(f.clone());
                }
            }
            84 => {
                if let (Some(f), Some(t)) = (cur.as_ref(), ext_text(rec)) {
                    out.push((f.face.clone(), f.bold, f.italic, t.0));
                }
            }
            _ => {}
        }
        off += size;
    }
    out
}

/// EXTTEXTOUTW: (text, reference x, y, dx per char in logical units)
fn ext_text(rec: &[u8]) -> Option<(String, i32, i32, Vec<i32>)> {
    let rx = rd_i32(rec, 36)?;
    let ry = rd_i32(rec, 40)?;
    let n = rd_u32(rec, 44)? as usize;
    let off_s = rd_u32(rec, 48)? as usize;
    let opts = rd_u32(rec, 52)?;
    let off_dx = rd_u32(rec, 72)? as usize;
    if n == 0 || n > 4096 {
        return None;
    }
    let units: Vec<u16> = (0..n)
        .filter_map(|i| rec.get(off_s + i * 2..off_s + i * 2 + 2))
        .map(|b| u16::from_le_bytes([b[0], b[1]]))
        .collect();
    let text = String::from_utf16_lossy(&units);
    // ETO_PDY (0x2000): dx and dy pairs
    let step = if opts & 0x2000 != 0 { 8 } else { 4 };
    let dx: Vec<i32> = if off_dx > 0 {
        (0..n).filter_map(|i| rd_i32(rec, off_dx + i * step)).collect()
    } else {
        Vec::new()
    };
    Some((text, rx, ry, dx))
}

#[derive(Clone)]
struct Dc {
    world: Xform,
    win_org: (f64, f64),
    win_ext: (f64, f64),
    vp_org: (f64, f64),
    vp_ext: (f64, f64),
    map_mode: u32,
    pen_color: Option<Color>,
    pen_width: f64,
    brush: Option<Color>,
    alternate: bool,
    cur: (f64, f64),
    font: Option<LogFont>,
    text_color: Color,
    text_align: u32,
}

fn rd_u32(b: &[u8], o: usize) -> Option<u32> {
    b.get(o..o + 4).map(|s| u32::from_le_bytes([s[0], s[1], s[2], s[3]]))
}
fn rd_i32(b: &[u8], o: usize) -> Option<i32> {
    rd_u32(b, o).map(|v| v as i32)
}
fn rd_i16(b: &[u8], o: usize) -> Option<i16> {
    b.get(o..o + 2).map(|s| i16::from_le_bytes([s[0], s[1]]))
}
fn rd_f32(b: &[u8], o: usize) -> Option<f32> {
    rd_u32(b, o).map(f32::from_bits)
}
fn colorref(v: u32) -> Color {
    Color::Rgb(
        (v & 0xFF) as f64 / 255.0,
        ((v >> 8) & 0xFF) as f64 / 255.0,
        ((v >> 16) & 0xFF) as f64 / 255.0,
    )
}

/// Is `data` an EMF (header record with the « EMF» signature)?
pub fn is_emf(data: &[u8]) -> bool {
    rd_u32(data, 0) == Some(1) && rd_u32(data, 40) == Some(0x464D_4520)
}

/// (face, bold, italic, char) -> PDF font key; (face, bold) -> (ascent, descent) per em
pub struct TextFaces<'f> {
    pub key: &'f dyn Fn(&str, bool, bool, char) -> String,
    pub metrics: &'f dyn Fn(&str, bool) -> (f64, f64),
}

struct Player<'a> {
    out: Vec<ContentElement>,
    dc: Dc,
    stack: Vec<Dc>,
    objs: std::collections::HashMap<u32, Obj>,
    path: Vec<PathOp>,
    in_path: bool,
    /// device units -> output points
    dev: Xform,
    dev_scale: f64,
    data: &'a [u8],
    faces: Option<&'a TextFaces<'a>>,
    mask: Option<((f64, f64, f64, f64), usize, usize, Vec<u8>)>,
}

impl<'a> Player<'a> {
    fn page(&self) -> Xform {
        let d = &self.dc;
        let (sx, sy) = match d.map_mode {
            // MM_ISOTROPIC / MM_ANISOTROPIC
            7 | 8 => (
                if d.win_ext.0 != 0.0 { d.vp_ext.0 / d.win_ext.0 } else { 1.0 },
                if d.win_ext.1 != 0.0 { d.vp_ext.1 / d.win_ext.1 } else { 1.0 },
            ),
            _ => (1.0, 1.0),
        };
        Xform {
            m11: sx,
            m12: 0.0,
            m21: 0.0,
            m22: sy,
            dx: d.vp_org.0 - d.win_org.0 * sx,
            dy: d.vp_org.1 - d.win_org.1 * sy,
        }
    }
    fn total(&self) -> Xform {
        self.dc.world.then(&self.page()).then(&self.dev)
    }
    fn pt(&self, x: f64, y: f64) -> (f64, f64) {
        self.total().apply(x, y)
    }
    fn width_pt(&self, w: f64) -> f64 {
        let t = self.total();
        let s = ((t.m11 * t.m11 + t.m12 * t.m12).sqrt() + (t.m21 * t.m21 + t.m22 * t.m22).sqrt()) / 2.0;
        if w <= 0.0 { self.dev_scale } else { (w * s).max(self.dev_scale * 0.5) }
    }
    fn stroke(&self) -> Option<StrokeStyle> {
        self.dc.pen_color.map(|c| StrokeStyle {
            color: c,
            width: self.width_pt(self.dc.pen_width),
            line_cap: LineCap::Round,
            line_join: LineJoin::Round,
        })
    }
    fn emit(&mut self, ops: Vec<PathOp>, fill: bool, stroke: bool) {
        if ops.is_empty() {
            return;
        }
        let fill_c = if fill { self.dc.brush } else { None };
        let stroke_s = if stroke { self.stroke() } else { None };
        if fill_c.is_none() && stroke_s.is_none() {
            return;
        }
        if let (Some(c), true) = (fill_c, self.dc.alternate) {
            // even-odd fill: clip by the path (W*), fill its bounding box
            let (mut x0, mut y0, mut x1, mut y1) = (f64::MAX, f64::MAX, f64::MIN, f64::MIN);
            for op in &ops {
                let pts: &[(f64, f64)] = &match op {
                    PathOp::MoveTo(x, y) | PathOp::LineTo(x, y) => vec![(*x, *y)],
                    PathOp::CurveTo(a, b, c2, d, e, f) => vec![(*a, *b), (*c2, *d), (*e, *f)],
                    PathOp::ClosePath => vec![],
                };
                for &(x, y) in pts {
                    x0 = x0.min(x); y0 = y0.min(y); x1 = x1.max(x); y1 = y1.max(y);
                }
            }
            if x0 <= x1 {
                self.out.push(ContentElement::SaveState);
                self.out.push(ContentElement::ClipPath(ClipPathData { operations: ops.clone(), even_odd: true }));
                self.out.push(ContentElement::Path(PathData {
                    operations: vec![
                        PathOp::MoveTo(x0, y0), PathOp::LineTo(x1, y0), PathOp::LineTo(x1, y1),
                        PathOp::LineTo(x0, y1), PathOp::ClosePath,
                    ],
                    stroke: None,
                    fill: Some(c),
                }));
                self.out.push(ContentElement::RestoreState);
            }
            if let Some(s) = stroke_s {
                self.out.push(ContentElement::Path(PathData { operations: ops, stroke: Some(s), fill: None }));
            }
        } else {
            self.out.push(ContentElement::Path(PathData { operations: ops, stroke: stroke_s, fill: fill_c }));
        }
    }
    fn points16(&self, rec: &[u8], at: usize, n: usize) -> Vec<(f64, f64)> {
        (0..n)
            .filter_map(|i| {
                let x = rd_i16(rec, at + i * 4)? as f64;
                let y = rd_i16(rec, at + i * 4 + 2)? as f64;
                Some(self.pt(x, y))
            })
            .collect()
    }
    fn points32(&self, rec: &[u8], at: usize, n: usize) -> Vec<(f64, f64)> {
        (0..n)
            .filter_map(|i| {
                let x = rd_i32(rec, at + i * 8)? as f64;
                let y = rd_i32(rec, at + i * 8 + 4)? as f64;
                Some(self.pt(x, y))
            })
            .collect()
    }
    fn poly_ops(pts: &[(f64, f64)], close: bool) -> Vec<PathOp> {
        let mut ops = Vec::with_capacity(pts.len() + 1);
        for (i, &(x, y)) in pts.iter().enumerate() {
            ops.push(if i == 0 { PathOp::MoveTo(x, y) } else { PathOp::LineTo(x, y) });
        }
        if close && !ops.is_empty() {
            ops.push(PathOp::ClosePath);
        }
        ops
    }
    fn bezier_ops(start: Option<(f64, f64)>, pts: &[(f64, f64)]) -> Vec<PathOp> {
        let mut ops = Vec::new();
        let rest = match start {
            Some(s) => {
                if !matches!(ops.last(), Some(_)) {
                    let _ = s;
                }
                pts
            }
            None => {
                if let Some(&(x, y)) = pts.first() {
                    ops.push(PathOp::MoveTo(x, y));
                }
                pts.get(1..).unwrap_or(&[])
            }
        };
        for c in rest.chunks_exact(3) {
            ops.push(PathOp::CurveTo(c[0].0, c[0].1, c[1].0, c[1].1, c[2].0, c[2].1));
        }
        ops
    }
    /// Append to the open path (inside BEGINPATH) or draw now.
    fn add(&mut self, ops: Vec<PathOp>, fill: bool) {
        if self.in_path {
            self.path.extend(ops);
        } else {
            self.emit(ops, fill, true);
        }
    }
    fn select(&mut self, ih: u32) {
        if ih & STOCK != 0 {
            match ih & !STOCK {
                0 => self.dc.brush = Some(Color::Gray(1.0)),
                1 => self.dc.brush = Some(Color::Gray(0.75)),
                2 => self.dc.brush = Some(Color::Gray(0.5)),
                3 => self.dc.brush = Some(Color::Gray(0.25)),
                4 => self.dc.brush = Some(Color::Gray(0.0)),
                5 => self.dc.brush = None,
                6 => { self.dc.pen_color = Some(Color::Gray(1.0)); self.dc.pen_width = 0.0 }
                7 => { self.dc.pen_color = Some(Color::Gray(0.0)); self.dc.pen_width = 0.0 }
                8 => self.dc.pen_color = None,
                _ => {}
            }
            return;
        }
        match self.objs.get(&ih).cloned() {
            Some(Obj::Pen { color, width }) => { self.dc.pen_color = color; self.dc.pen_width = width }
            Some(Obj::Brush(c)) => self.dc.brush = c,
            Some(Obj::Font(f)) => self.dc.font = Some(f),
            None => {}
        }
    }
    fn decode_dib(rec: &[u8], off_bmi: usize, off_bits: usize) -> Option<(usize, usize, Vec<u8>)> {
        let bi = rec.get(off_bmi..)?;
        let (Some(hsz), Some(w), Some(h), Some(bpp), Some(comp)) = (
            rd_u32(bi, 0),
            rd_i32(bi, 4),
            rd_i32(bi, 8),
            bi.get(14..16).map(|s| u16::from_le_bytes([s[0], s[1]])),
            rd_u32(bi, 16),
        ) else { return None };
        if comp != 0 || w <= 0 || h == 0 || w > 20000 || h.abs() > 20000 {
            return None;
        }
        let (wu, hu) = (w as usize, h.unsigned_abs() as usize);
        let clr_used = rd_u32(bi, 32).unwrap_or(0) as usize;
        let ncol = if bpp <= 8 { if clr_used > 0 { clr_used } else { 1 << bpp } } else { 0 };
        let pal: Vec<[u8; 3]> = (0..ncol)
            .filter_map(|i| bi.get(hsz as usize + i * 4..hsz as usize + i * 4 + 3).map(|c| [c[2], c[1], c[0]]))
            .collect();
        let stride = ((wu * bpp as usize + 31) / 32) * 4;
        let bits = rec.get(off_bits..)?;
        let mut rgb = Vec::with_capacity(wu * hu * 3);
        for row in 0..hu {
            let src_row = if h > 0 { hu - 1 - row } else { row };
            let line = bits.get(src_row * stride..src_row * stride + stride)?;
            for x in 0..wu {
                let px = match bpp {
                    32 => line.get(x * 4..x * 4 + 3).map(|c| [c[2], c[1], c[0]]),
                    24 => line.get(x * 3..x * 3 + 3).map(|c| [c[2], c[1], c[0]]),
                    // BI_RGB 16 bpp = X1R5G5B5
                    16 => line.get(x * 2..x * 2 + 2).map(|c| {
                        let v = u16::from_le_bytes([c[0], c[1]]);
                        let e = |k: u16| ((k & 31) as u32 * 255 / 31) as u8;
                        [e(v >> 10), e(v >> 5), e(v)]
                    }),
                    8 => line.get(x).and_then(|&i| pal.get(i as usize).copied()),
                    4 => line.get(x / 2).and_then(|&b| pal.get(if x % 2 == 0 { b >> 4 } else { b & 15 } as usize).copied()),
                    1 => line.get(x / 8).and_then(|&b| pal.get(((b >> (7 - x % 8)) & 1) as usize).copied()),
                    _ => None,
                };
                rgb.extend_from_slice(&px?);
            }
        }
        Some((wu, hu, rgb))
    }
    /// A raster record: SRCCOPY draws; SRCPAINT (OR) holds a mask whose white
    /// pixels the next SRCAND image at the same place leaves transparent -- the
    /// classic two-pass transparent bitmap (ja administrative__1fc69a4f's clip-art);
    /// other raster operations are not drawn.
    fn dib(&mut self, rec: &[u8], off_bmi: usize, off_bits: usize, dest: [(f64, f64); 2], rop: u32) {
        let Some((wu, hu, rgb)) = Self::decode_dib(rec, off_bmi, off_bits) else { return };
        let (x0, y0) = (dest[0].0.min(dest[1].0), dest[0].1.min(dest[1].1));
        let (x1, y1) = (dest[0].0.max(dest[1].0), dest[0].1.max(dest[1].1));
        let key = (x0, y0, x1, y1);
        let alpha = match rop {
            0x00EE_0086 => {
                self.mask = Some((key, wu, hu, rgb));
                return;
            }
            0x0088_00C6 => match self.mask.take() {
                Some((k, mw, mh, m)) if mw == wu && mh == hu
                    && (k.0 - x0).abs() < 0.5 && (k.1 - y0).abs() < 0.5 && (k.2 - x1).abs() < 0.5 && (k.3 - y1).abs() < 0.5 =>
                {
                    Some(m.chunks(3).map(|c| 255 - ((c[0] as u32 + c[1] as u32 + c[2] as u32) / 3) as u8).collect())
                }
                _ => None,
            },
            0x00CC_0020 => None,
            _ => return,
        };
        self.out.push(ContentElement::Image(ImageData {
            x: x0,
            y: y0,
            width: x1 - x0,
            height: y1 - y0,
            data: rgb,
            color_space: ColorSpace::DeviceRgb,
            bits_per_component: 8,
            pixel_width: wu as u32,
            pixel_height: hu as u32,
            alpha,
        }));
        let _ = self.data;
    }
    fn record(&mut self, kind: u32, rec: &[u8]) {
        match kind {
            9 => if let (Some(x), Some(y)) = (rd_i32(rec, 8), rd_i32(rec, 12)) { self.dc.win_ext = (x as f64, y as f64) },
            10 => if let (Some(x), Some(y)) = (rd_i32(rec, 8), rd_i32(rec, 12)) { self.dc.win_org = (x as f64, y as f64) },
            11 => if let (Some(x), Some(y)) = (rd_i32(rec, 8), rd_i32(rec, 12)) { self.dc.vp_ext = (x as f64, y as f64) },
            12 => if let (Some(x), Some(y)) = (rd_i32(rec, 8), rd_i32(rec, 12)) { self.dc.vp_org = (x as f64, y as f64) },
            17 => if let Some(m) = rd_u32(rec, 8) { self.dc.map_mode = m },
            19 => if let Some(m) = rd_u32(rec, 8) { self.dc.alternate = m == 1 },
            33 => self.stack.push(self.dc.clone()),
            34 => {
                let n = rd_i32(rec, 8).unwrap_or(-1);
                let k = if n < 0 { (-n) as usize } else { 1 };
                for _ in 0..k {
                    if let Some(d) = self.stack.pop() { self.dc = d; }
                }
            }
            35 | 36 => {
                let f = |o| rd_f32(rec, o).unwrap_or(0.0) as f64;
                let x = Xform { m11: f(8), m12: f(12), m21: f(16), m22: f(20), dx: f(24), dy: f(28) };
                let mode = if kind == 35 { 4 } else { rd_u32(rec, 32).unwrap_or(4) };
                self.dc.world = match mode {
                    1 => Xform::ID,
                    2 => x.then(&self.dc.world),
                    3 => self.dc.world.then(&x),
                    _ => x,
                };
            }
            37 => if let Some(ih) = rd_u32(rec, 8) { self.select(ih) },
            38 => {
                if let (Some(ih), Some(style), Some(w), Some(c)) = (rd_u32(rec, 8), rd_u32(rec, 12), rd_i32(rec, 16), rd_u32(rec, 24)) {
                    let color = if style & 0xF == 5 { None } else { Some(colorref(c)) };
                    self.objs.insert(ih, Obj::Pen { color, width: w as f64 });
                }
            }
            95 => {
                if let (Some(ih), Some(style), Some(w), Some(c)) = (rd_u32(rec, 8), rd_u32(rec, 28), rd_u32(rec, 32), rd_u32(rec, 40)) {
                    let color = if style & 0xF == 5 { None } else { Some(colorref(c)) };
                    self.objs.insert(ih, Obj::Pen { color, width: w as f64 });
                }
            }
            39 => {
                if let (Some(ih), Some(style), Some(c)) = (rd_u32(rec, 8), rd_u32(rec, 12), rd_u32(rec, 16)) {
                    self.objs.insert(ih, Obj::Brush(if style == 1 { None } else { Some(colorref(c)) }));
                }
            }
            40 => if let Some(ih) = rd_u32(rec, 8) { self.objs.remove(&ih); },
            82 => if let (Some(ih), Some(f)) = (rd_u32(rec, 8), logfont(rec)) { self.objs.insert(ih, Obj::Font(f)); },
            24 => if let Some(c) = rd_u32(rec, 8) { self.dc.text_color = colorref(c) },
            22 => if let Some(a) = rd_u32(rec, 8) { self.dc.text_align = a },
            84 => self.text(rec),
            27 => if let (Some(x), Some(y)) = (rd_i32(rec, 8), rd_i32(rec, 12)) {
                let p = self.pt(x as f64, y as f64);
                self.dc.cur = p;
                if self.in_path { self.path.push(PathOp::MoveTo(p.0, p.1)); }
            },
            54 => if let (Some(x), Some(y)) = (rd_i32(rec, 8), rd_i32(rec, 12)) {
                let p = self.pt(x as f64, y as f64);
                if self.in_path {
                    self.path.push(PathOp::LineTo(p.0, p.1));
                } else {
                    let c = self.dc.cur;
                    self.emit(vec![PathOp::MoveTo(c.0, c.1), PathOp::LineTo(p.0, p.1)], false, true);
                }
                self.dc.cur = p;
            },
            59 => { self.in_path = true; self.path.clear(); }
            60 => self.in_path = false,
            61 => self.path.push(PathOp::ClosePath),
            68 => { self.in_path = false; self.path.clear(); }
            62 | 63 | 64 => {
                let ops = std::mem::take(&mut self.path);
                self.in_path = false;
                self.emit(ops, kind != 64, kind != 62);
            }
            // 16-bit point lists: bounds 8..24, count at 24, points from 28
            85..=89 => {
                let n = rd_u32(rec, 24).unwrap_or(0) as usize;
                let pts = self.points16(rec, 28, n);
                self.poly(kind - 83, pts);
            }
            // 32-bit: POLYBEZIER 2, POLYGON 3, POLYLINE 4, POLYBEZIERTO 5, POLYLINETO 6
            2..=6 => {
                let n = rd_u32(rec, 24).unwrap_or(0) as usize;
                let pts = self.points32(rec, 28, n);
                self.poly(kind, pts);
            }
            // POLYPOLYGON16 91 / POLYPOLYLINE16 90 / POLYPOLYGON 8 / POLYPOLYLINE 7
            7 | 8 | 90 | 91 => {
                let np = rd_u32(rec, 24).unwrap_or(0) as usize;
                let total = rd_u32(rec, 28).unwrap_or(0) as usize;
                let counts: Vec<usize> = (0..np).filter_map(|i| rd_u32(rec, 32 + i * 4).map(|v| v as usize)).collect();
                let start = 32 + np * 4;
                let all = if kind >= 90 { self.points16(rec, start, total) } else { self.points32(rec, start, total) };
                let close = kind == 8 || kind == 91;
                let mut ops = Vec::new();
                let mut at = 0;
                for c in counts {
                    if let Some(seg) = all.get(at..at + c) { ops.extend(Self::poly_ops(seg, close)); }
                    at += c;
                }
                self.add(ops, close);
            }
            42 | 43 => {
                if let (Some(l), Some(t), Some(r), Some(b)) = (rd_i32(rec, 8), rd_i32(rec, 12), rd_i32(rec, 16), rd_i32(rec, 20)) {
                    let (l, t, r, b) = (l as f64, t as f64, r as f64, b as f64);
                    let ops = if kind == 43 {
                        let p = [self.pt(l, t), self.pt(r, t), self.pt(r, b), self.pt(l, b)];
                        Self::poly_ops(&p, true)
                    } else {
                        let (cx, cy, rx, ry) = ((l + r) / 2.0, (t + b) / 2.0, (r - l) / 2.0, (b - t) / 2.0);
                        let k = 0.5523;
                        let q = |x: f64, y: f64| self.pt(x, y);
                        let (p0, p1, p2, p3) = (q(cx + rx, cy), q(cx, cy + ry), q(cx - rx, cy), q(cx, cy - ry));
                        let c = |a: (f64, f64), b: (f64, f64), e: (f64, f64)| PathOp::CurveTo(a.0, a.1, b.0, b.1, e.0, e.1);
                        vec![
                            PathOp::MoveTo(p0.0, p0.1),
                            c(q(cx + rx, cy + ry * k), q(cx + rx * k, cy + ry), p1),
                            c(q(cx - rx * k, cy + ry), q(cx - rx, cy + ry * k), p2),
                            c(q(cx - rx, cy - ry * k), q(cx - rx * k, cy - ry), p3),
                            c(q(cx + rx * k, cy - ry), q(cx + rx, cy - ry * k), p0),
                            PathOp::ClosePath,
                        ]
                    };
                    self.add(ops, true);
                }
            }
            // BITBLT: with no source bitmap it paints the destination with the
            // brush (PATCOPY) or white / black (WHITENESS / BLACKNESS) -- a chart's
            // filled cells; with one, the bitmap.
            76 => {
                if let (Some(xd), Some(yd), Some(cx), Some(cy), Some(rop), Some(ob), Some(cb), Some(obits)) = (
                    rd_i32(rec, 24), rd_i32(rec, 28), rd_i32(rec, 32), rd_i32(rec, 36), rd_u32(rec, 40),
                    rd_u32(rec, 84), rd_u32(rec, 88), rd_u32(rec, 92),
                ) {
                    let (l, t, r, b) = (xd as f64, yd as f64, (xd + cx) as f64, (yd + cy) as f64);
                    if cb == 0 {
                        let fill = match rop {
                            0x00F0_0021 => self.dc.brush,
                            0x00FF_0062 => Some(Color::Gray(1.0)),
                            0x0000_0042 => Some(Color::Gray(0.0)),
                            _ => None,
                        };
                        if let Some(c) = fill {
                            let q = [self.pt(l, t), self.pt(r, t), self.pt(r, b), self.pt(l, b)];
                            self.out.push(ContentElement::Path(PathData {
                                operations: Self::poly_ops(&q, true),
                                stroke: None,
                                fill: Some(c),
                            }));
                        }
                    } else {
                        let a = self.pt(l, t);
                        let z = self.pt(r, b);
                        self.dib(rec, ob as usize, obits as usize, [a, z], rop);
                    }
                }
            }
            // STRETCHDIBITS
            81 => {
                if let (Some(xd), Some(yd), Some(ob), Some(obits), Some(cxd), Some(cyd), Some(rop)) = (
                    rd_i32(rec, 24), rd_i32(rec, 28), rd_u32(rec, 48), rd_u32(rec, 56), rd_i32(rec, 72), rd_i32(rec, 76), rd_u32(rec, 68),
                ) {
                    let a = self.pt(xd as f64, yd as f64);
                    let b = self.pt((xd + cxd) as f64, (yd + cyd) as f64);
                    self.dib(rec, ob as usize, obits as usize, [a, b], rop);
                }
            }
            _ => {}
        }
    }
    fn text(&mut self, rec: &[u8]) {
        let (Some(faces), Some(f)) = (self.faces, self.dc.font.clone()) else { return };
        // rotated text is not set yet: drawn level it would cross the picture
        if f.escapement % 3600 != 0 {
            return;
        }
        let Some((text, rx, ry, dx)) = ext_text(rec) else { return };
        let t = self.total();
        let yscale = (t.m21 * t.m21 + t.m22 * t.m22).sqrt();
        let (asc, desc) = (faces.metrics)(&f.face, f.bold);
        let em = if f.height < 0 { -f.height as f64 } else { f.height as f64 / (asc + desc).max(0.5) };
        let size = em * yscale;
        if size < 0.5 {
            return;
        }
        let chars: Vec<char> = text.chars().collect();
        let total_dx: f64 = dx.iter().map(|&d| d as f64).sum();
        // TA_RIGHT 2, TA_CENTER 6
        let shift = match self.dc.text_align & 6 {
            6 => -total_dx / 2.0,
            2 => -total_dx,
            _ => 0.0,
        };
        // TA_BASELINE 24, TA_BOTTOM 8, TA_TOP 0; logical y may run up the page
        let down = if t.m22 >= 0.0 { 1.0 } else { -1.0 };
        let base_dy = match self.dc.text_align & 24 {
            24 => 0.0,
            8 => -desc * em * down,
            _ => asc * em * down,
        };
        let mut ax = rx as f64 + shift;
        for (i, &c) in chars.iter().enumerate() {
            if !c.is_whitespace() {
                let (px, py) = self.pt(ax, ry as f64 + base_dy);
                self.out.push(ContentElement::Text(TextSpan {
                    x: px,
                    y: py,
                    text: c.to_string(),
                    font_name: (faces.key)(&f.face, f.bold, f.italic, c),
                    font_size: size,
                    fill_color: self.dc.text_color,
                    character_spacing: 0.0,
                    stroke_width: 0.0,
                    skew: 0.0,
                    horizontal_scale: 100.0,
                    stroke_color: None,
                }));
            }
            ax += dx.get(i).copied().unwrap_or(0) as f64;
        }
    }
    fn poly(&mut self, kind: u32, pts: Vec<(f64, f64)>) {
        match kind {
            // POLYBEZIER
            2 => self.add(Self::bezier_ops(None, &pts), false),
            // POLYGON
            3 => self.add(Self::poly_ops(&pts, true), true),
            // POLYLINE
            4 => self.add(Self::poly_ops(&pts, false), false),
            // POLYBEZIERTO
            5 => {
                let mut ops = Vec::new();
                if !self.in_path { ops.push(PathOp::MoveTo(self.dc.cur.0, self.dc.cur.1)); }
                ops.extend(Self::bezier_ops(Some(self.dc.cur), &pts));
                if let Some(&l) = pts.last() { self.dc.cur = l; }
                self.add(ops, false);
            }
            // POLYLINETO
            6 => {
                let mut ops = Vec::new();
                if !self.in_path { ops.push(PathOp::MoveTo(self.dc.cur.0, self.dc.cur.1)); }
                ops.extend(pts.iter().map(|&(x, y)| PathOp::LineTo(x, y)));
                if let Some(&l) = pts.last() { self.dc.cur = l; }
                self.add(ops, false);
            }
            _ => {}
        }
    }
}

/// Play `data` into the box (x, y, w, h) (points, y down). `None` when it is
/// not an EMF or draws nothing.
pub fn render(data: &[u8], x: f64, y: f64, w: f64, h: f64, faces: Option<&TextFaces>,
    crop: Option<(f32, f32, f32, f32)>) -> Option<Vec<ContentElement>> {
    if !is_emf(data) || std::env::var_os("OXI_PDF_EMF_DISABLE").is_some() {
        return None;
    }
    // frame (0.01 mm) and the reference device (pixels <-> millimetres)
    let fl = rd_i32(data, 24)? as f64;
    let ft = rd_i32(data, 28)? as f64;
    let fr = rd_i32(data, 32)? as f64;
    let fb = rd_i32(data, 36)? as f64;
    let dev_px = (rd_i32(data, 72)? as f64, rd_i32(data, 76)? as f64);
    let dev_mm = (rd_i32(data, 80)? as f64, rd_i32(data, 84)? as f64);
    // a:srcRect (percent insets top, right, bottom, left) narrows the frame that
    // fills the box (reports__00d20c7b: r 20.0 b 20.1 -- the chart drew 0.8x small)
    let (fl, ft, fr, fb) = match crop {
        Some((t, r, b, l)) => {
            let (fw, fh) = (fr - fl, fb - ft);
            (fl + fw * l as f64 / 100.0, ft + fh * t as f64 / 100.0, fr - fw * r as f64 / 100.0, fb - fh * b as f64 / 100.0)
        }
        None => (fl, ft, fr, fb),
    };
    if fr <= fl || fb <= ft || dev_px.0 <= 0.0 || dev_px.1 <= 0.0 || dev_mm.0 <= 0.0 || dev_mm.1 <= 0.0 {
        return None;
    }
    // device px -> 0.01 mm -> box
    let sx = dev_mm.0 * 100.0 / dev_px.0 * w / (fr - fl);
    let sy = dev_mm.1 * 100.0 / dev_px.1 * h / (fb - ft);
    let dev = Xform {
        m11: sx,
        m12: 0.0,
        m21: 0.0,
        m22: sy,
        dx: x - fl * w / (fr - fl),
        dy: y - ft * h / (fb - ft),
    };
    let mut p = Player {
        out: vec![ContentElement::SaveState, ContentElement::ClipPath(ClipPathData {
            operations: vec![
                PathOp::MoveTo(x, y), PathOp::LineTo(x + w, y), PathOp::LineTo(x + w, y + h),
                PathOp::LineTo(x, y + h), PathOp::ClosePath,
            ],
            even_odd: false,
        })],
        dc: Dc {
            world: Xform::ID,
            win_org: (0.0, 0.0),
            win_ext: (1.0, 1.0),
            vp_org: (0.0, 0.0),
            vp_ext: (1.0, 1.0),
            map_mode: 1,
            pen_color: Some(Color::Gray(0.0)),
            pen_width: 0.0,
            brush: Some(Color::Gray(1.0)),
            alternate: true,
            cur: (0.0, 0.0),
            font: None,
            text_color: Color::Gray(0.0),
            text_align: 0,
        },
        stack: Vec::new(),
        objs: std::collections::HashMap::new(),
        path: Vec::new(),
        in_path: false,
        dev,
        dev_scale: sx.abs().max(sy.abs()),
        data,
        faces,
        mask: None,
    };
    let start_len = p.out.len();
    let mut off = 0usize;
    while off + 8 <= data.len() {
        let kind = rd_u32(data, off)?;
        let size = rd_u32(data, off + 4)? as usize;
        if size < 8 || off + size > data.len() {
            break;
        }
        if kind == 14 {
            break;
        }
        let rec = &data[off..off + size];
        p.record(kind, rec);
        off += size;
    }
    if p.out.len() == start_len {
        return None;
    }
    p.out.push(ContentElement::RestoreState);
    Some(p.out)
}
