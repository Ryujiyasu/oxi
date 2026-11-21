// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! `.pptx` to PDF pages: one page per slide, in the same page model
//! `docx-to-pdf` writes, so `write_pdf` and the rasteriser serve both.
//!
//! The text is laid out by `oxislides_core::layout` -- the rules measured
//! from PowerPoint -- with advances read from the faces this machine has
//! where the compiled tables do not cover a family. What this module adds is
//! the drawing: backgrounds, shape outlines, pictures, tables and the glyphs.
//!
//! What it does not draw is reported, not skipped silently: charts, gradient
//! fills, rotated text and pictures, text warps and preset shapes with no
//! outline here are listed as [`UnsupportedContent`], which `--strict` turns
//! into a refusal.

use std::cell::RefCell;
use std::collections::{BTreeMap, BTreeSet, HashMap};
use std::rc::Rc;

use oxidocs_core::ir::UnsupportedContent;
use oxipdf_core::ir::*;
use oxislides_core::ir::{
    CellBorder, GeomCmd, MasterStyleLevel, Presentation, Shape, ShapeContent, Slide, SlideBullet,
    SlideParagraph, Table,
};
use oxislides_core::layout::{self, FaceMetrics, PlacedLine, TableMetrics};
use rustybuzz::ttf_parser;

/// One font file (or collection member) this conversion draws with.
struct FaceFile {
    path: String,
    index: u32,
    data: Vec<u8>,
    /// The name its glyphs are embedded under.
    key: String,
    advances: RefCell<HashMap<char, Option<f32>>>,
    /// `1.2 * ascent / (ascent + descent)`, the baseline rule the renderer
    /// asks its font system for.
    baseline_em: Option<f32>,
}

impl FaceFile {
    fn load(path: &str, index: u32, key: String) -> Option<FaceFile> {
        let data = std::fs::read(path).ok()?;
        let baseline_em = {
            let face = ttf_parser::Face::parse(&data, index).ok()?;
            let (asc, desc) = (f32::from(face.ascender()), -f32::from(face.descender()));
            (asc + desc > 0.0).then(|| 1.2 * asc / (asc + desc))
        };
        Some(FaceFile {
            path: path.to_string(),
            index,
            data,
            key,
            advances: RefCell::new(HashMap::new()),
            baseline_em,
        })
    }

    /// The advance of `ch` in em, or None when the face has no glyph for it.
    fn advance_em(&self, ch: char) -> Option<f32> {
        if let Some(hit) = self.advances.borrow().get(&ch) {
            return *hit;
        }
        let found = ttf_parser::Face::parse(&self.data, self.index).ok().and_then(|face| {
            let glyph = face.glyph_index(ch)?;
            let units = f32::from(face.units_per_em());
            Some(f32::from(face.glyph_hor_advance(glyph)?) / units)
        });
        self.advances.borrow_mut().insert(ch, found);
        found
    }
}

/// Lookups already made, including the ones that found nothing.
type Cached<K> = RefCell<HashMap<K, Option<Rc<FaceFile>>>>;

/// The faces of one conversion: which file serves a family, and which file
/// draws a character its family's file lacks.
#[derive(Default)]
struct Faces {
    by_request: Cached<(String, bool, bool)>,
    by_file: Cached<(String, u32)>,
    fallbacks: RefCell<Option<Vec<Rc<FaceFile>>>>,
    missing: RefCell<BTreeSet<String>>,
}

impl Faces {
    fn file(&self, path: &str, index: u32) -> Option<Rc<FaceFile>> {
        let id = (path.to_string(), index);
        if let Some(hit) = self.by_file.borrow().get(&id) {
            return hit.clone();
        }
        let key = format!("OxiSlide-{}", self.by_file.borrow().len());
        let loaded = FaceFile::load(path, index, key).map(Rc::new);
        self.by_file.borrow_mut().insert(id, loaded.clone());
        loaded
    }

    /// The file for a family at a weight and slope: the family itself, its
    /// metric-compatible stand-in, the same pair with a style word taken off
    /// the name, and last a sans face so the text is drawn at all.
    fn primary(&self, family: &str, bold: bool, italic: bool) -> Option<Rc<FaceFile>> {
        let request = (family.trim().to_lowercase(), bold, italic);
        if let Some(hit) = self.by_request.borrow().get(&request) {
            return hit.clone();
        }
        let index = super::font_index();
        let lookup = |name: &str, bold: bool, italic: bool| {
            index
                .find(name, bold, italic)
                .or_else(|| super::substitute_for(name).and_then(|s| index.find(s, bold, italic)))
                .map(|face| (face.path.to_string_lossy().to_string(), face.index))
        };
        let found = lookup(family, bold, italic).or_else(|| {
            super::split_style_suffix(family).and_then(|(base, b, i)| lookup(&base, bold || b, italic || i))
        });
        let found = found.or_else(|| {
            self.missing.borrow_mut().insert(family.trim().to_string());
            ["Liberation Sans", "Carlito", "Arial", "DejaVu Sans"]
                .iter()
                .find_map(|stand_in| lookup(stand_in, bold, italic))
        });
        let face = found.and_then(|(path, index)| self.file(&path, index));
        self.by_request.borrow_mut().insert(request, face.clone());
        face
    }

    /// Faces tried, in order, for a character the family's own file lacks:
    /// an East Asian face, then the symbol faces `docx-to-pdf` falls back to.
    fn fallbacks(&self) -> Vec<Rc<FaceFile>> {
        if let Some(list) = self.fallbacks.borrow().as_ref() {
            return list.clone();
        }
        let east_asian = super::CJK_FONT_PATHS
            .iter()
            .map(|p| p.to_string())
            .find(|p| std::path::Path::new(p).exists())
            .or_else(|| super::bundled_face_file("OxiGothic.ttf").map(|p| p.to_string_lossy().to_string()));
        let list: Vec<Rc<FaceFile>> = east_asian
            .into_iter()
            .chain(super::FALLBACK_FACES.iter().map(|p| p.to_string()))
            .filter(|p| std::path::Path::new(p).exists())
            .filter_map(|p| self.file(&p, 0))
            .collect();
        *self.fallbacks.borrow_mut() = Some(list.clone());
        list
    }

    /// The face that draws `ch` for this request and its advance in em.
    fn face_for(&self, family: &str, bold: bool, italic: bool, ch: char) -> Option<(Rc<FaceFile>, f32)> {
        let primary = self.primary(family, bold, italic);
        if let Some(face) = &primary {
            if let Some(advance) = face.advance_em(ch) {
                return Some((face.clone(), advance));
            }
        }
        for face in self.fallbacks() {
            if let Some(advance) = face.advance_em(ch) {
                return Some((face, advance));
            }
        }
        // No face has it: it is set as the family's missing glyph, at a width
        // that keeps the line from collapsing onto it.
        primary.map(|face| (face, 0.5))
    }
}

/// The compiled PowerPoint tables where they cover a family, this machine's
/// font files where they do not.
struct Metrics<'a> {
    faces: &'a Faces,
}

impl FaceMetrics for Metrics<'_> {
    fn advance_em(&self, family: &str, bold: bool, italic: bool, ch: char) -> Option<f32> {
        TableMetrics
            .advance_em(family, bold, italic, ch)
            .or_else(|| self.faces.face_for(family, bold, italic, ch).map(|(_, advance)| advance))
    }

    fn has_all_glyphs(&self, family: &str, bold: bool, italic: bool, text: &str) -> bool {
        text.chars().all(|c| self.advance_em(family, bold, italic, c).is_some())
    }

    fn baseline_offset_em(&self, family: &str) -> Option<f32> {
        TableMetrics
            .baseline_offset_em(family)
            .or_else(|| self.faces.primary(family, false, false).and_then(|face| face.baseline_em))
    }
}

/// What the slides hold that this module does not draw, by kind and slide.
#[derive(Default)]
struct Census {
    seen: BTreeMap<(String, String), usize>,
}

impl Census {
    fn note(&mut self, slide: usize, kind: &str) {
        *self
            .seen
            .entry((kind.to_string(), format!("ppt/slides/slide{}.xml", slide + 1)))
            .or_default() += 1;
    }

    fn into_list(self) -> Vec<UnsupportedContent> {
        self.seen
            .into_iter()
            .map(|((kind, part), count)| UnsupportedContent { kind, part, count })
            .collect()
    }
}

/// A piece of text placed on a page, before its face has been embedded.
struct Glyphs {
    face: Rc<FaceFile>,
    x: f32,
    y: f32,
    text: String,
    size: f32,
    color: Color,
    /// The face is not the weight or slope that was asked for, so the PDF
    /// thickens or slants it.
    fake_bold: bool,
    fake_italic: bool,
}

enum Item {
    Element(ContentElement),
    Glyphs(Glyphs),
}

/// Box-local points to page points: flip inside the box, turn about its
/// centre, move to its place.
#[derive(Clone, Copy)]
struct Placement {
    x: f32,
    y: f32,
    w: f32,
    h: f32,
    flip_h: bool,
    flip_v: bool,
    sin: f32,
    cos: f32,
}

impl Placement {
    fn of(shape: &Shape) -> Placement {
        let turn = shape.rotation.to_radians();
        Placement {
            x: shape.x,
            y: shape.y,
            w: shape.width,
            h: shape.height,
            flip_h: shape.flip_h,
            flip_v: shape.flip_v,
            sin: turn.sin(),
            cos: turn.cos(),
        }
    }

    fn point(&self, x: f32, y: f32) -> (f64, f64) {
        let x = if self.flip_h { self.w - x } else { x };
        let y = if self.flip_v { self.h - y } else { y };
        let (dx, dy) = (x - self.w / 2.0, y - self.h / 2.0);
        (
            f64::from(self.x + self.w / 2.0 + dx * self.cos - dy * self.sin),
            f64::from(self.y + self.h / 2.0 + dx * self.sin + dy * self.cos),
        )
    }
}

/// An outline in box-local points.
#[derive(Default)]
struct Outline {
    ops: Vec<GeomCmd>,
    /// A line has no inside to fill.
    open: bool,
}

impl Outline {
    fn polygon(points: &[(f32, f32)]) -> Outline {
        let mut ops: Vec<GeomCmd> = points
            .iter()
            .enumerate()
            .map(|(i, &(x, y))| if i == 0 { GeomCmd::MoveTo(x, y) } else { GeomCmd::LineTo(x, y) })
            .collect();
        ops.push(GeomCmd::Close);
        Outline { ops, open: false }
    }

    fn rect(w: f32, h: f32) -> Outline {
        Outline::polygon(&[(0.0, 0.0), (w, 0.0), (w, h), (0.0, h)])
    }

    fn ellipse(w: f32, h: f32) -> Outline {
        Outline::round_rect(w, h, w / 2.0, h / 2.0)
    }

    /// A rectangle with quarter-ellipse corners of radii `rx`, `ry`.
    fn round_rect(w: f32, h: f32, rx: f32, ry: f32) -> Outline {
        // The cubic that stands in for a quarter circle.
        const K: f32 = 0.552_284_8;
        let (rx, ry) = (rx.min(w / 2.0), ry.min(h / 2.0));
        let (kx, ky) = (rx * (1.0 - K), ry * (1.0 - K));
        Outline {
            ops: vec![
                GeomCmd::MoveTo(rx, 0.0),
                GeomCmd::LineTo(w - rx, 0.0),
                GeomCmd::CubicTo(w - kx, 0.0, w, ky, w, ry),
                GeomCmd::LineTo(w, h - ry),
                GeomCmd::CubicTo(w, h - ky, w - kx, h, w - rx, h),
                GeomCmd::LineTo(rx, h),
                GeomCmd::CubicTo(kx, h, 0.0, h - ky, 0.0, h - ry),
                GeomCmd::LineTo(0.0, ry),
                GeomCmd::CubicTo(0.0, ky, kx, 0.0, rx, 0.0),
                GeomCmd::Close,
            ],
            open: false,
        }
    }

    /// Cubics along the ellipse about (`cx`, `cy`) with radii `rx`, `ry`,
    /// from parametric angle `from` through `sweep` radians, a quarter turn
    /// at most per cubic. The pen is assumed to be at the start already.
    fn arc(ops: &mut Vec<GeomCmd>, (cx, cy): (f32, f32), (rx, ry): (f32, f32), from: f32, sweep: f32) {
        let steps = (sweep.abs() / std::f32::consts::FRAC_PI_2).ceil().max(1.0) as usize;
        let step = sweep / steps as f32;
        // Handle length for a step of the unit circle.
        let k = 4.0 / 3.0 * (step / 4.0).tan();
        for i in 0..steps {
            let (a, b) = (from + step * i as f32, from + step * (i + 1) as f32);
            let (sa, ca, sb, cb) = (a.sin(), a.cos(), b.sin(), b.cos());
            ops.push(GeomCmd::CubicTo(
                cx + rx * (ca - k * sa),
                cy + ry * (sa + k * ca),
                cx + rx * (cb + k * sb),
                cy + ry * (sb - k * cb),
                cx + rx * cb,
                cy + ry * sb,
            ));
        }
    }

    /// `wedgeEllipseCallout` (ECMA-376 presetShapeDefinitions): the ellipse,
    /// opened 11 degrees either side of the direction of the point, with the
    /// two ends joined to the point.
    fn ellipse_callout(w: f32, h: f32, tip: (f32, f32)) -> Outline {
        let (rx, ry) = (w / 2.0, h / 2.0);
        let (dx, dy) = (tip.0 - rx, tip.1 - ry);
        // The direction of the point in the ellipse's own parameter: the
        // spec's `at2 (dxPos * h) (dyPos * w)`.
        let toward = (dy * w).atan2(dx * h);
        let gap = 11f32.to_radians();
        let from = toward + gap;
        let mut ops = vec![GeomCmd::MoveTo(rx + rx * from.cos(), ry + ry * from.sin())];
        Outline::arc(&mut ops, (rx, ry), (rx, ry), from, std::f32::consts::TAU - 2.0 * gap);
        ops.push(GeomCmd::LineTo(tip.0, tip.1));
        ops.push(GeomCmd::Close);
        Outline { ops, open: false }
    }

    /// `wedgeRectCallout` and `wedgeRoundRectCallout`: the box, with a wedge
    /// to the point leaving the side the point is furthest past, between the
    /// 2/12 and 5/12 marks of that side (7/12 and 10/12 when the point is on
    /// the far half). `r` is the corner radius, 0 for the square one.
    fn rect_callout(w: f32, h: f32, tip: (f32, f32), r: f32) -> Outline {
        let (dx_pos, dy_pos) = (tip.0 - w / 2.0, tip.1 - h / 2.0);
        // Which way the point leaves: across or up/down, with the box squared.
        let across = (dx_pos * h / w).abs() > dy_pos.abs();
        let (x1, x2) = if dx_pos > 0.0 { (w * 7.0 / 12.0, w * 10.0 / 12.0) } else { (w * 2.0 / 12.0, w * 5.0 / 12.0) };
        let (y1, y2) = if dy_pos > 0.0 { (h * 7.0 / 12.0, h * 10.0 / 12.0) } else { (h * 2.0 / 12.0, h * 5.0 / 12.0) };
        let r = r.clamp(0.0, w.min(h) / 2.0);
        let wedge_top = !across && dy_pos <= 0.0;
        let wedge_right = across && dx_pos > 0.0;
        let wedge_bottom = !across && dy_pos > 0.0;
        let wedge_left = across && dx_pos <= 0.0;
        const K: f32 = 0.552_284_8;
        let mut ops = vec![GeomCmd::MoveTo(r, 0.0)];
        if wedge_top {
            ops.extend([GeomCmd::LineTo(x1, 0.0), GeomCmd::LineTo(tip.0, tip.1), GeomCmd::LineTo(x2, 0.0)]);
        }
        ops.push(GeomCmd::LineTo(w - r, 0.0));
        if r > 0.0 {
            ops.push(GeomCmd::CubicTo(w - r + r * K, 0.0, w, r - r * K, w, r));
        }
        if wedge_right {
            ops.extend([GeomCmd::LineTo(w, y1), GeomCmd::LineTo(tip.0, tip.1), GeomCmd::LineTo(w, y2)]);
        }
        ops.push(GeomCmd::LineTo(w, h - r));
        if r > 0.0 {
            ops.push(GeomCmd::CubicTo(w, h - r + r * K, w - r + r * K, h, w - r, h));
        }
        if wedge_bottom {
            ops.extend([GeomCmd::LineTo(x2, h), GeomCmd::LineTo(tip.0, tip.1), GeomCmd::LineTo(x1, h)]);
        }
        ops.push(GeomCmd::LineTo(r, h));
        if r > 0.0 {
            ops.push(GeomCmd::CubicTo(r - r * K, h, 0.0, h - r + r * K, 0.0, h - r));
        }
        if wedge_left {
            ops.extend([GeomCmd::LineTo(0.0, y2), GeomCmd::LineTo(tip.0, tip.1), GeomCmd::LineTo(0.0, y1)]);
        }
        ops.push(GeomCmd::LineTo(0.0, r));
        if r > 0.0 {
            ops.push(GeomCmd::CubicTo(0.0, r - r * K, r - r * K, 0.0, r, 0.0));
        }
        ops.push(GeomCmd::Close);
        Outline { ops, open: false }
    }

    /// The outline's on-curve points on the page, in drawing order.
    fn points(&self, at: &Placement) -> Vec<(f64, f64)> {
        self.ops
            .iter()
            .filter_map(|op| match *op {
                GeomCmd::MoveTo(x, y) | GeomCmd::LineTo(x, y) | GeomCmd::CubicTo(_, _, _, _, x, y) => {
                    Some(at.point(x, y))
                }
                GeomCmd::Close => None,
            })
            .collect()
    }

    fn path_ops(&self, at: &Placement) -> Vec<PathOp> {
        self.ops
            .iter()
            .map(|op| match *op {
                GeomCmd::MoveTo(x, y) => {
                    let (x, y) = at.point(x, y);
                    PathOp::MoveTo(x, y)
                }
                GeomCmd::LineTo(x, y) => {
                    let (x, y) = at.point(x, y);
                    PathOp::LineTo(x, y)
                }
                GeomCmd::CubicTo(x1, y1, x2, y2, x, y) => {
                    let (x1, y1) = at.point(x1, y1);
                    let (x2, y2) = at.point(x2, y2);
                    let (x, y) = at.point(x, y);
                    PathOp::CurveTo(x1, y1, x2, y2, x, y)
                }
                GeomCmd::Close => PathOp::ClosePath,
            })
            .collect()
    }
}

/// A preset's adjustment as a fraction, or its default.
fn adjust(shape: &Shape, name: &str, default: f32) -> f32 {
    shape.adjustments.get(name).copied().unwrap_or(default) / 100_000.0
}

/// How a part of a shape is filled: not at all, with the shape's fill, or
/// with the fill darkened (DrawingML's `darkenLess` on a folded corner).
#[derive(Clone, Copy, PartialEq)]
enum PartFill {
    None,
    Shape,
    DarkenLess,
}

/// `darkenLess` scales each channel to 80%.
const DARKEN_LESS: f64 = 0.8;

/// The outlines of a shape, in box-local points, each with how it is
/// filled. None when the preset has no outline here.
fn outlines(shape: &Shape) -> Option<Vec<(Outline, PartFill)>> {
    let (w, h) = (shape.width, shape.height);
    if let Some(geometry) = shape.custom_geometry.as_ref().filter(|g| !g.unsupported) {
        let paths = geometry
            .paths
            .iter()
            .filter(|path| !path.commands.is_empty())
            .map(|path| {
                // Each path is drawn in its own space, stretched over the box.
                let (sx, sy) = (
                    if path.w > 0.0 { w / path.w } else { 1.0 },
                    if path.h > 0.0 { h / path.h } else { 1.0 },
                );
                let ops = path
                    .commands
                    .iter()
                    .map(|op| match *op {
                        GeomCmd::MoveTo(x, y) => GeomCmd::MoveTo(x * sx, y * sy),
                        GeomCmd::LineTo(x, y) => GeomCmd::LineTo(x * sx, y * sy),
                        GeomCmd::CubicTo(a, b, c, d, x, y) => {
                            GeomCmd::CubicTo(a * sx, b * sy, c * sx, d * sy, x * sx, y * sy)
                        }
                        GeomCmd::Close => GeomCmd::Close,
                    })
                    .collect();
                let fill = if path.fill_none { PartFill::None } else { PartFill::Shape };
                (Outline { ops, open: false }, fill)
            })
            .collect::<Vec<_>>();
        if !paths.is_empty() {
            return Some(paths);
        }
    }
    let short = w.min(h);
    let one = |outline: Outline| Some(vec![(outline, PartFill::Shape)]);
    match shape.shape_type.as_deref().unwrap_or("rect") {
        "rect" | "flowChartProcess" => one(Outline::rect(w, h)),
        "ellipse" | "flowChartConnector" => one(Outline::ellipse(w, h)),
        "roundRect" | "flowChartAlternateProcess" => {
            let r = short * adjust(shape, "adj", 16_667.0).clamp(0.0, 0.5);
            one(Outline::round_rect(w, h, r, r))
        }
        "triangle" => {
            let apex = w * adjust(shape, "adj", 50_000.0).clamp(0.0, 1.0);
            one(Outline::polygon(&[(apex, 0.0), (w, h), (0.0, h)]))
        }
        "rtTriangle" => one(Outline::polygon(&[(0.0, 0.0), (w, h), (0.0, h)])),
        "diamond" | "flowChartDecision" => {
            one(Outline::polygon(&[(w / 2.0, 0.0), (w, h / 2.0), (w / 2.0, h), (0.0, h / 2.0)]))
        }
        "parallelogram" => {
            let lean = (short * adjust(shape, "adj", 25_000.0)).clamp(0.0, w);
            one(Outline::polygon(&[(lean, 0.0), (w, 0.0), (w - lean, h), (0.0, h)]))
        }
        "trapezoid" => {
            let lean = (short * adjust(shape, "adj", 25_000.0)).clamp(0.0, w / 2.0);
            one(Outline::polygon(&[(lean, 0.0), (w - lean, 0.0), (w, h), (0.0, h)]))
        }
        "pentagon" | "homePlate" => {
            let point = (short * adjust(shape, "adj", 50_000.0)).clamp(0.0, w);
            one(Outline::polygon(&[(0.0, 0.0), (w - point, 0.0), (w, h / 2.0), (w - point, h), (0.0, h)]))
        }
        "chevron" => {
            let point = (short * adjust(shape, "adj", 50_000.0)).clamp(0.0, w);
            one(Outline::polygon(&[
                (0.0, 0.0),
                (w - point, 0.0),
                (w, h / 2.0),
                (w - point, h),
                (0.0, h),
                (point, h / 2.0),
            ]))
        }
        "hexagon" => {
            let cut = (short * adjust(shape, "adj", 25_000.0)).clamp(0.0, w / 2.0);
            one(Outline::polygon(&[
                (cut, 0.0),
                (w - cut, 0.0),
                (w, h / 2.0),
                (w - cut, h),
                (cut, h),
                (0.0, h / 2.0),
            ]))
        }
        "rightArrow" | "leftArrow" => {
            let shaft = h * adjust(shape, "adj1", 50_000.0).clamp(0.0, 1.0);
            let head = (short * adjust(shape, "adj2", 50_000.0)).clamp(0.0, w);
            let (top, bottom) = ((h - shaft) / 2.0, (h + shaft) / 2.0);
            let right = [
                (0.0, top),
                (w - head, top),
                (w - head, 0.0),
                (w, h / 2.0),
                (w - head, h),
                (w - head, bottom),
                (0.0, bottom),
            ];
            if shape.shape_type.as_deref() == Some("leftArrow") {
                let left: Vec<(f32, f32)> = right.iter().map(|&(x, y)| (w - x, y)).collect();
                one(Outline::polygon(&left))
            } else {
                one(Outline::polygon(&right))
            }
        }
        // A folded corner (ECMA-376 presetShapeDefinitions): the page with its
        // bottom-right corner cut off at `dy2`, and the fold drawn over the
        // cut in the fill darkened.
        "foldedCorner" => {
            let dy2 = short * adjust(shape, "adj", 16_667.0).clamp(0.0, 0.5);
            let dy10 = dy2 / 5.0;
            let (x1, y2) = (w - dy2, h - dy2);
            Some(vec![
                (Outline::polygon(&[(0.0, 0.0), (w, 0.0), (w, y2), (x1, h), (0.0, h)]), PartFill::Shape),
                (
                    Outline::polygon(&[(x1, h), (x1 + dy10, y2 + dy10), (w, y2)]),
                    PartFill::DarkenLess,
                ),
            ])
        }
        // The point of a callout is placed by adj1 / adj2, as fractions of
        // the width and height from the centre.
        "wedgeEllipseCallout" | "wedgeRectCallout" | "wedgeRoundRectCallout" => {
            let tip = (
                w / 2.0 + w * adjust(shape, "adj1", -20_833.0),
                h / 2.0 + h * adjust(shape, "adj2", 62_500.0),
            );
            match shape.shape_type.as_deref() {
                Some("wedgeEllipseCallout") => one(Outline::ellipse_callout(w, h, tip)),
                Some("wedgeRectCallout") => one(Outline::rect_callout(w, h, tip, 0.0)),
                _ => one(Outline::rect_callout(w, h, tip, short * adjust(shape, "adj3", 16_667.0))),
            }
        }
        "line" | "straightConnector1" => Some(vec![(
            Outline { ops: vec![GeomCmd::MoveTo(0.0, 0.0), GeomCmd::LineTo(w, h)], open: true },
            PartFill::None,
        )]),
        // An elbow: across to `adj1` of the width, down the full height, then
        // across to the end. The renderer's S-BENTCONN, which matched
        // PowerPoint's own PDF on all four points to 0.01pt.
        "bentConnector3" => {
            let xb = w * adjust(shape, "adj1", 50_000.0);
            Some(vec![(
                Outline {
                    ops: vec![
                        GeomCmd::MoveTo(0.0, 0.0),
                        GeomCmd::LineTo(xb, 0.0),
                        GeomCmd::LineTo(xb, h),
                        GeomCmd::LineTo(w, h),
                    ],
                    open: true,
                },
                PartFill::None,
            )])
        }
        _ => None,
    }
}

/// The filled outline of a line end at `tip`, pointing away from `behind`,
/// for a line `width` points wide.
///
/// The geometry is the renderer's, read off PowerPoint's own PDF: the size
/// tokens scale the line width, floored at 2pt, by 2 / 3 / 5 for sm / med /
/// lg; a triangle's tip is on the end, a stealth head is notched two thirds
/// of the way back, and an oval and a diamond are centred on the end.
///
/// `Some(None)` for `none`; None for a kind with no outline here (`arrow`,
/// an open V PowerPoint strokes, which the renderer leaves undrawn too).
fn line_end(
    end: &oxislides_core::ir::LineEnd,
    tip: (f64, f64),
    behind: (f64, f64),
    width: f64,
) -> Option<Option<Vec<PathOp>>> {
    let factor = |token: &str| match token {
        "sm" => 2.0,
        "lg" => 5.0,
        _ => 3.0,
    };
    let (dx, dy) = (tip.0 - behind.0, tip.1 - behind.1);
    let length = (dx * dx + dy * dy).sqrt();
    if length < 1e-6 {
        return Some(None);
    }
    let (ux, uy) = (dx / length, dy / length);
    let (vx, vy) = (-uy, ux);
    let unit = width.max(2.0);
    let across = factor(&end.w) * unit;
    let along = factor(&end.len) * unit;
    let at = |u: f64, v: f64| (tip.0 + u * ux + v * vx, tip.1 + u * uy + v * vy);
    let corners: Vec<(f64, f64)> = match end.kind.as_str() {
        "none" => return Some(None),
        "triangle" => vec![at(0.0, 0.0), at(-along, across / 2.0), at(-along, -across / 2.0)],
        "stealth" => vec![
            at(0.0, 0.0),
            at(-along, across / 2.0),
            at(-along * 2.0 / 3.0, 0.0),
            at(-along, -across / 2.0),
        ],
        "diamond" => vec![
            at(along / 2.0, 0.0),
            at(0.0, across / 2.0),
            at(-along / 2.0, 0.0),
            at(0.0, -across / 2.0),
        ],
        "oval" => (0..32)
            .map(|i| {
                let t = f64::from(i) * std::f64::consts::TAU / 32.0;
                at(along / 2.0 * t.cos(), across / 2.0 * t.sin())
            })
            .collect(),
        _ => return None,
    };
    let mut operations: Vec<PathOp> = corners
        .iter()
        .enumerate()
        .map(|(i, &(x, y))| if i == 0 { PathOp::MoveTo(x, y) } else { PathOp::LineTo(x, y) })
        .collect();
    operations.push(PathOp::ClosePath);
    Some(Some(operations))
}

/// `color` at `alpha` over white. The page model has no transparency, so a
/// translucent fill is drawn as the opaque colour it shows on a white slide.
fn over_white(color: Color, alpha: f32) -> Color {
    let a = f64::from(alpha.clamp(0.0, 1.0));
    match color {
        Color::Rgb(r, g, b) => Color::Rgb(r * a + (1.0 - a), g * a + (1.0 - a), b * a + (1.0 - a)),
        other => other,
    }
}

fn solid(hex: Option<&str>, alpha: Option<f32>) -> Option<Color> {
    let color = super::parse_hex_color(hex?)?;
    Some(match alpha {
        Some(a) if a < 1.0 => over_white(color, a),
        _ => color,
    })
}

/// Which outline levels a shape's text inherits, and its theme face.
fn text_context<'a>(pres: &'a Presentation, shape: &Shape) -> (&'a [MasterStyleLevel], &'a str) {
    match shape.ph_type.as_deref() {
        Some("title") | Some("ctrTitle") => (&pres.master_styles.title, &pres.major_font),
        Some(_) => (&pres.master_styles.body, &pres.minor_font),
        None => (&pres.master_styles.other, &pres.minor_font),
    }
}

/// The run a character of a paragraph belongs to.
fn run_at(paragraph: &SlideParagraph, char_index: usize) -> Option<&oxislides_core::ir::SlideRun> {
    let mut seen = 0;
    for run in &paragraph.runs {
        seen += run.text.chars().count();
        if char_index < seen {
            return Some(run);
        }
    }
    paragraph.runs.last()
}

/// The marker a numbered paragraph shows for its position in the list.
fn auto_number(kind: &str, n: u32) -> String {
    let letters = |n: u32, base: u8| {
        let mut n = n;
        let mut out = Vec::new();
        while n > 0 {
            n -= 1;
            out.push(base + (n % 26) as u8);
            n /= 26;
        }
        out.reverse();
        String::from_utf8(out).unwrap_or_default()
    };
    let roman = |mut n: u32| {
        let mut out = String::new();
        for (value, digits) in [
            (1000, "m"), (900, "cm"), (500, "d"), (400, "cd"), (100, "c"), (90, "xc"),
            (50, "l"), (40, "xl"), (10, "x"), (9, "ix"), (5, "v"), (4, "iv"), (1, "i"),
        ] {
            while n >= value {
                out.push_str(digits);
                n -= value;
            }
        }
        out
    };
    let body = if kind.starts_with("alphaLc") {
        letters(n, b'a')
    } else if kind.starts_with("alphaUc") {
        letters(n, b'A')
    } else if kind.starts_with("romanLc") {
        roman(n)
    } else if kind.starts_with("romanUc") {
        roman(n).to_uppercase()
    } else {
        n.to_string()
    };
    if kind.ends_with("ParenBoth") {
        format!("({body})")
    } else if kind.ends_with("ParenR") {
        format!("{body})")
    } else if kind.ends_with("Period") {
        format!("{body}.")
    } else {
        body
    }
}

struct SlideDraw<'a> {
    pres: &'a Presentation,
    faces: &'a Faces,
    census: &'a mut Census,
    slide: usize,
    items: Vec<Item>,
}

impl SlideDraw<'_> {
    fn push(&mut self, element: ContentElement) {
        self.items.push(Item::Element(element));
    }

    fn path(&mut self, operations: Vec<PathOp>, fill: Option<Color>, stroke: Option<StrokeStyle>) {
        if fill.is_some() || stroke.is_some() {
            self.push(ContentElement::Path(PathData { operations, stroke, fill }));
        }
    }

    fn background(&mut self, slide: &Slide) {
        let (w, h) = (self.pres.slide_width, self.pres.slide_height);
        let whole = Placement { x: 0.0, y: 0.0, w, h, flip_h: false, flip_v: false, sin: 0.0, cos: 1.0 };
        if let Some(color) = solid(slide.background_color.as_deref(), None) {
            self.path(Outline::rect(w, h).path_ops(&whole), Some(color), None);
        }
        if let Some(gradient) = &slide.background_gradient {
            // Drawn as its first stop, so the slide is not left white, and
            // reported: a gradient shown as one colour is not the slide.
            self.census.note(self.slide, "background-gradient");
            let first = gradient.stops.first().and_then(|stop| solid(Some(&stop.color), Some(stop.alpha)));
            self.path(Outline::rect(w, h).path_ops(&whole), first, None);
        }
        if let Some(image) = &slide.background_image {
            self.image(&image.data, 0.0, 0.0, w, h, None);
        }
    }

    /// Draw encoded image bytes into a box, cropped to `crop` (left, top,
    /// right, bottom as fractions of the source).
    fn image(&mut self, data: &[u8], x: f32, y: f32, w: f32, h: f32, crop: Option<(f32, f32, f32, f32)>) {
        let Ok(decoded) = image::load_from_memory(data) else {
            self.census.note(self.slide, "picture-format");
            return;
        };
        let decoded = match crop {
            Some((l, t, r, b)) if [l, t, r, b].iter().all(|v| (0.0..1.0).contains(v)) && l + r < 1.0 && t + b < 1.0 => {
                let (pw, ph) = (decoded.width() as f32, decoded.height() as f32);
                decoded.crop_imm(
                    (pw * l).round() as u32,
                    (ph * t).round() as u32,
                    ((pw * (1.0 - l - r)).round() as u32).max(1),
                    ((ph * (1.0 - t - b)).round() as u32).max(1),
                )
            }
            _ => decoded,
        };
        let alpha = decoded.color().has_alpha().then(|| {
            decoded.to_rgba8().pixels().map(|pixel| pixel.0[3]).collect::<Vec<u8>>()
        });
        let alpha = alpha.filter(|a| a.iter().any(|&v| v < 255));
        let rgb = decoded.to_rgb8();
        let (pixel_width, pixel_height) = rgb.dimensions();
        self.push(ContentElement::Image(ImageData {
            x: f64::from(x),
            y: f64::from(y),
            width: f64::from(w),
            height: f64::from(h),
            data: rgb.into_raw(),
            color_space: ColorSpace::DeviceRgb,
            bits_per_component: 8,
            pixel_width,
            pixel_height,
            alpha,
        }));
    }

    fn shape(&mut self, shape: &Shape) {
        let at = Placement::of(shape);
        let turned = shape.rotation.rem_euclid(360.0).abs() > 0.01 || shape.flip_h || shape.flip_v;

        if let ShapeContent::Image { data, .. } = &shape.content {
            if turned {
                self.census.note(self.slide, "rotated-picture");
            }
            // The destination a `fillRect` asks for, inside the shape's box.
            let (mut x, mut y, mut w, mut h) = (shape.x, shape.y, shape.width, shape.height);
            if let Some((l, t, r, b)) = shape.fill_rect {
                x += shape.width * l;
                y += shape.height * t;
                w -= shape.width * (l + r);
                h -= shape.height * (t + b);
                if l < 0.0 || t < 0.0 || r < 0.0 || b < 0.0 {
                    // The part past the box is not shown.
                    self.push(ContentElement::SaveState);
                    self.push(ContentElement::ClipPath(ClipPathData {
                        operations: Outline::rect(shape.width, shape.height).path_ops(&at),
                        even_odd: false,
                    }));
                    self.image(data, x, y, w, h, shape.src_rect);
                    self.push(ContentElement::RestoreState);
                    self.outline_stroke(shape, &at);
                    return;
                }
            }
            self.image(data, x, y, w, h, shape.src_rect);
            self.outline_stroke(shape, &at);
            return;
        }

        self.fill_and_stroke(shape, &at);

        match &shape.content {
            ShapeContent::AutoShape { paragraphs } | ShapeContent::TextBox { paragraphs } => {
                if paragraphs.iter().any(|p| p.runs.iter().any(|r| !r.text.trim().is_empty())) {
                    if turned {
                        self.census.note(self.slide, "rotated-text");
                    }
                    if shape.text_warp.as_deref().is_some_and(|warp| warp != "textNoShape") {
                        self.census.note(self.slide, "text-warp");
                    }
                }
                self.text(shape, paragraphs);
            }
            ShapeContent::Table { table } => self.table(shape, table),
            ShapeContent::Chart { .. } => self.census.note(self.slide, "chart"),
            ShapeContent::Unsupported { element_type } => {
                let kind = format!("element:{element_type}");
                self.census.note(self.slide, &kind);
            }
            ShapeContent::Placeholder | ShapeContent::Image { .. } => {}
        }
    }

    fn stroke_of(shape: &Shape) -> Option<StrokeStyle> {
        let color = solid(shape.border_color.as_deref(), shape.border_alpha)?;
        let width = f64::from(shape.border_width.unwrap_or(0.75));
        (width > 0.0).then_some(StrokeStyle {
            color,
            width,
            line_cap: match shape.line_cap.as_deref() {
                Some("rnd") => LineCap::Round,
                Some("sq") => LineCap::Square,
                _ => LineCap::Butt,
            },
            line_join: LineJoin::Miter,
        })
    }

    /// A picture's own border, if it has one.
    fn outline_stroke(&mut self, shape: &Shape, at: &Placement) {
        if let Some(stroke) = Self::stroke_of(shape) {
            self.path(Outline::rect(shape.width, shape.height).path_ops(at), None, Some(stroke));
        }
    }

    fn fill_and_stroke(&mut self, shape: &Shape, at: &Placement) {
        let mut fill = solid(shape.fill_color.as_deref(), shape.fill_alpha);
        if let Some(gradient) = &shape.gradient {
            self.census.note(self.slide, "gradient-fill");
            fill = gradient.stops.first().and_then(|stop| solid(Some(&stop.color), Some(stop.alpha)));
        }
        let stroke = Self::stroke_of(shape);
        if fill.is_none() && stroke.is_none() {
            return;
        }
        if shape.border_dash.as_deref().is_some_and(|dash| dash != "solid") {
            // The page model has no dash pattern; the line is drawn solid.
            self.census.note(self.slide, "dashed-line");
        }
        let Some(parts) = outlines(shape) else {
            let kind = format!("preset:{}", shape.shape_type.as_deref().unwrap_or("?"));
            self.census.note(self.slide, &kind);
            return;
        };
        for (outline, part) in &parts {
            let fill = match (*part, fill, outline.open) {
                (_, _, true) | (PartFill::None, ..) => None,
                (PartFill::Shape, fill, _) => fill,
                (PartFill::DarkenLess, fill, _) => fill.map(|color| match color {
                    Color::Rgb(r, g, b) => Color::Rgb(r * DARKEN_LESS, g * DARKEN_LESS, b * DARKEN_LESS),
                    other => other,
                }),
            };
            self.path(outline.path_ops(at), fill, stroke.clone());
        }

        // Line ends sit on an open outline's first and last points, pointing
        // along its first and last segments.
        let Some(stroke) = stroke else { return };
        let Some(points) = parts.iter().find(|(outline, _)| outline.open).map(|(outline, _)| outline.points(at)) else {
            if shape.head_end.is_some() || shape.tail_end.is_some() {
                self.census.note(self.slide, "line-end");
            }
            return;
        };
        if points.len() < 2 {
            return;
        }
        let n = points.len();
        for (end, tip, behind) in [
            (&shape.head_end, points[0], points[1]),
            (&shape.tail_end, points[n - 1], points[n - 2]),
        ] {
            let Some(end) = end else { continue };
            match line_end(end, tip, behind, stroke.width) {
                Some(Some(operations)) => self.path(operations, Some(stroke.color), None),
                Some(None) => {}
                None => {
                    let kind = format!("line-end:{}", end.kind);
                    self.census.note(self.slide, &kind);
                }
            }
        }
    }

    /// Lay out and place a text body. `shape` gives the box, its insets and
    /// its anchor; the box is drawn unturned.
    fn text(&mut self, shape: &Shape, paragraphs: &[SlideParagraph]) {
        if paragraphs.is_empty() {
            return;
        }
        let (master, theme_face) = text_context(self.pres, shape);
        let metrics = Metrics { faces: self.faces };
        let placed = layout::layout_text_shape(&metrics, shape, paragraphs, master, &shape.ph_levels, theme_face);
        // Numbering runs per level and restarts when a shallower level
        // intervenes.
        let mut counters: HashMap<u32, u32> = HashMap::new();
        let mut marked: BTreeSet<usize> = BTreeSet::new();
        for line in &placed.lines {
            let Some(paragraph) = paragraphs.get(line.para_index) else { continue };
            let level = layout::resolve_level(master, &shape.ph_levels, paragraph.lvl);
            if marked.insert(line.para_index) {
                self.marker(shape, paragraph, &level, line, &mut counters);
            }
            self.line(shape, paragraph, &level, line, &metrics);
        }
    }

    /// The bullet or number at the head of a paragraph's first line.
    fn marker(
        &mut self,
        shape: &Shape,
        paragraph: &SlideParagraph,
        level: &MasterStyleLevel,
        line: &PlacedLine,
        counters: &mut HashMap<u32, u32>,
    ) {
        let bullet = match &paragraph.bullet {
            SlideBullet::Inherit => &level.bullet,
            own => own,
        };
        let lvl = paragraph.lvl;
        let (text, face) = match bullet {
            SlideBullet::Char { ch, font, .. } => {
                counters.retain(|&l, _| l < lvl);
                (ch.to_string(), font.clone())
            }
            SlideBullet::AutoNum { kind, start_at } => {
                counters.retain(|&l, _| l <= lvl);
                let n = counters.entry(lvl).or_insert(start_at.unwrap_or(1).saturating_sub(1));
                *n += 1;
                (auto_number(kind, *n), None)
            }
            SlideBullet::Inherit | SlideBullet::None => {
                counters.retain(|&l, _| l < lvl);
                return;
            }
        };
        if line.text.trim().is_empty() {
            // PowerPoint shows no marker on an empty paragraph.
            return;
        }
        let (geom_l, ..) =
            layout::geom_text_insets(shape.shape_type.as_deref(), &shape.adjustments, shape.width, shape.height);
        let geometry = layout::indent_geometry(
            shape.width,
            paragraph.mar_l.unwrap_or(level.mar_l),
            paragraph.indent.unwrap_or(level.indent),
            false,
        );
        let color = run_at(paragraph, 0)
            .and_then(|run| solid(run.color.as_deref(), run.color_alpha))
            .or_else(|| solid(level.color.as_deref(), None))
            .unwrap_or(Color::Rgb(0.0, 0.0, 0.0));
        // A symbol font's codes are the Unicode symbols they draw wherever
        // the symbol font itself is not installed.
        let family = face.unwrap_or_else(|| line.family.clone());
        let text: String = text
            .chars()
            .map(|c| oxidocs_fonts::symbol_font_to_unicode(&family, c).unwrap_or(c))
            .collect();
        let mut x = shape.x + shape.l_ins + geom_l + geometry.marker_x;
        for ch in text.chars() {
            let Some((face, advance)) = self.faces.face_for(&family, false, false, ch) else { continue };
            self.items.push(Item::Glyphs(Glyphs {
                face,
                x,
                y: shape.y + line.baseline,
                text: ch.to_string(),
                size: line.font_size,
                color,
                fake_bold: false,
                fake_italic: false,
            }));
            x += advance * line.font_size;
        }
    }

    fn line(
        &mut self,
        shape: &Shape,
        paragraph: &SlideParagraph,
        level: &MasterStyleLevel,
        line: &PlacedLine,
        metrics: &Metrics,
    ) {
        // Where the engine put each character; where it could not say, the
        // faces' own advances.
        let offsets = layout::glyph_offsets_pt(metrics, line);
        let whole = [layout::LineSegment {
            text: line.text.clone(),
            family: line.family.clone(),
            font_size: line.font_size,
            bold: line.bold,
            italic: line.italic,
            start: 0,
        }];
        let segments = if line.segments.is_empty() { &whole[..] } else { &line.segments[..] };
        let default_color = solid(level.color.as_deref(), None).unwrap_or(Color::Rgb(0.0, 0.0, 0.0));

        let mut pen = 0.0f32;
        let mut index = 0usize;
        let mut open: Option<Glyphs> = None;
        let mut underlines: Vec<(f32, f32, f32, Color)> = Vec::new();
        for segment in segments {
            for ch in segment.text.chars() {
                let at = offsets.as_ref().and_then(|o| o.get(index).copied()).unwrap_or(pen);
                let found = self.faces.face_for(&segment.family, segment.bold, segment.italic, ch);
                let advance = found.as_ref().map_or(0.5, |(_, advance)| *advance) * segment.font_size;
                let run = run_at(paragraph, line.char_start + index);
                let color = run
                    .and_then(|run| solid(run.color.as_deref(), run.color_alpha))
                    .unwrap_or(default_color);
                let end = offsets.as_ref().and_then(|o| o.get(index + 1).copied()).unwrap_or(at + advance);
                if run.is_some_and(|run| run.underline) {
                    underlines.push((at, end, segment.font_size, color));
                }
                index += 1;
                pen = at + advance;
                let Some((face, _)) = found else { continue };
                if ch.is_whitespace() || (ch as u32) < 0x20 {
                    // A word is placed where the engine put its first letter,
                    // so a face whose advances differ cannot drift the line.
                    self.items.extend(open.take().map(Item::Glyphs));
                    continue;
                }
                let continues = open.as_ref().is_some_and(|glyphs| {
                    Rc::ptr_eq(&glyphs.face, &face)
                        && (glyphs.size - segment.font_size).abs() < 0.001
                        && format!("{:?}", glyphs.color) == format!("{color:?}")
                });
                if continues {
                    if let Some(glyphs) = open.as_mut() {
                        glyphs.text.push(ch);
                    }
                    continue;
                }
                self.items.extend(open.take().map(Item::Glyphs));
                let served = self.faces.primary(&segment.family, segment.bold, segment.italic);
                let own_style = served.as_ref().is_some_and(|primary| Rc::ptr_eq(primary, &face));
                let index_face = super::font_index().find(&segment.family, segment.bold, segment.italic);
                open = Some(Glyphs {
                    face,
                    x: shape.x + line.x + at,
                    y: shape.y + line.baseline,
                    text: ch.to_string(),
                    size: segment.font_size,
                    color,
                    // Only a fallback face, which has one weight, is faked.
                    fake_bold: segment.bold && !own_style && index_face.is_none(),
                    fake_italic: segment.italic && !own_style && index_face.is_none(),
                });
            }
        }
        self.items.extend(open.take().map(Item::Glyphs));

        for (from, to, size, color) in underlines {
            let y = f64::from(shape.y + line.baseline + size * 0.1);
            self.path(
                vec![
                    PathOp::MoveTo(f64::from(shape.x + line.x + from), y),
                    PathOp::LineTo(f64::from(shape.x + line.x + to), y),
                ],
                None,
                Some(StrokeStyle {
                    color,
                    width: f64::from(size * 0.05),
                    line_cap: LineCap::Butt,
                    line_join: LineJoin::Miter,
                }),
            );
        }
    }

    fn table(&mut self, shape: &Shape, table: &Table) {
        if table.col_widths.is_empty() {
            return;
        }
        // A cell's text is laid out as a text box the size of the cell.
        let cell_box = |x: f32, y: f32, w: f32, h: f32, cell: &oxislides_core::ir::TableCell| {
            let mut frame = shape.clone();
            frame.x = x;
            frame.y = y;
            frame.width = w;
            frame.height = h;
            frame.rotation = 0.0;
            frame.flip_h = false;
            frame.flip_v = false;
            frame.shape_type = None;
            frame.ph_type = None;
            frame.ph_levels = Vec::new();
            frame.custom_geometry = None;
            frame.l_ins = cell.mar_l;
            frame.r_ins = cell.mar_r;
            frame.t_ins = cell.mar_t;
            frame.b_ins = cell.mar_b;
            frame.anchor = cell.anchor.clone();
            frame.wrap_text = true;
            frame.content = ShapeContent::Placeholder;
            frame
        };
        let span_width = |col: usize, span: u32| -> f32 {
            table.col_widths.iter().skip(col).take(span.max(1) as usize).sum()
        };
        let metrics = Metrics { faces: self.faces };
        let master = &self.pres.master_styles.other[..];

        // A row is as tall as it says, or as tall as its fullest cell needs.
        let mut heights: Vec<f32> = Vec::with_capacity(table.rows.len());
        for (r, row) in table.rows.iter().enumerate() {
            let mut height = table.row_heights.get(r).copied().unwrap_or(0.0);
            let mut col = 0usize;
            for cell in row {
                let span = cell.grid_span.max(1);
                if !cell.h_merge && !cell.paragraphs.is_empty() {
                    let frame = cell_box(0.0, 0.0, span_width(col, span), height, cell);
                    let placed = layout::layout_text_shape(
                        &metrics, &frame, &cell.paragraphs, master, &[], &self.pres.minor_font,
                    );
                    height = height.max(placed.height + cell.mar_t + cell.mar_b);
                }
                col += if cell.h_merge { 0 } else { span as usize };
            }
            heights.push(height);
        }

        let mut y = shape.y;
        for (r, row) in table.rows.iter().enumerate() {
            let height = heights[r];
            let mut x = shape.x;
            let mut col = 0usize;
            for cell in row {
                if cell.h_merge {
                    continue;
                }
                let span = cell.grid_span.max(1);
                let width = span_width(col, span);
                let frame = cell_box(x, y, width, height, cell);
                let at = Placement::of(&frame);
                if let Some(fill) = solid(cell.fill_color.as_deref(), cell.fill_alpha) {
                    self.path(Outline::rect(width, height).path_ops(&at), Some(fill), None);
                }
                // Left, right, top, bottom.
                let edges = [
                    ((0.0, 0.0), (0.0, height)),
                    ((width, 0.0), (width, height)),
                    ((0.0, 0.0), (width, 0.0)),
                    ((0.0, height), (width, height)),
                ];
                for (border, (from, to)) in cell.borders.iter().zip(edges) {
                    let Some(CellBorder { color, width: line, alpha }) = border else { continue };
                    let Some(color) = solid(Some(color), Some(*alpha)) else { continue };
                    if *line <= 0.0 {
                        continue;
                    }
                    let (x1, y1) = at.point(from.0, from.1);
                    let (x2, y2) = at.point(to.0, to.1);
                    self.path(
                        vec![PathOp::MoveTo(x1, y1), PathOp::LineTo(x2, y2)],
                        None,
                        Some(StrokeStyle {
                            color,
                            width: f64::from(*line),
                            line_cap: LineCap::Butt,
                            line_join: LineJoin::Miter,
                        }),
                    );
                }
                self.text(&frame, &cell.paragraphs);
                x += width;
                col += span as usize;
            }
            y += height;
        }
    }
}

/// Embed each face with the characters it draws and turn the placed text
/// into spans that name it.
fn finish(pages: Vec<(f64, f64, Vec<Item>)>, faces: &Faces) -> PdfDocument {
    let mut wanted: BTreeMap<String, (Rc<FaceFile>, BTreeSet<char>)> = BTreeMap::new();
    for (_, _, items) in &pages {
        for item in items {
            if let Item::Glyphs(glyphs) = item {
                wanted
                    .entry(glyphs.face.key.clone())
                    .or_insert_with(|| (glyphs.face.clone(), BTreeSet::new()))
                    .1
                    .extend(glyphs.text.chars());
            }
        }
    }
    let mut embedded_fonts = HashMap::new();
    for (key, (face, chars)) in wanted {
        let font = if face.index == 0 {
            super::subset_font_file(&face.path, &chars, &key)
        } else {
            super::embed_collection_member(&face.path, face.index, &key)
        };
        match font {
            Some(font) => {
                embedded_fonts.insert(key, font);
            }
            None => eprintln!("Warning: could not embed {} - its text is drawn by a viewer substitute", face.path),
        }
    }
    let missing = faces.missing.borrow();
    if !missing.is_empty() {
        let list: Vec<&str> = missing.iter().map(String::as_str).collect();
        eprintln!(
            "Warning: no font file for {} - drawn with a stand-in face and may differ from the source",
            list.join(", ")
        );
    }

    let pages = pages
        .into_iter()
        .map(|(width, height, items)| Page {
            width,
            height,
            media_box: Rectangle { llx: 0.0, lly: 0.0, urx: width, ury: height },
            crop_box: None,
            rotation: 0,
            contents: items
                .into_iter()
                .map(|item| match item {
                    Item::Element(element) => element,
                    Item::Glyphs(glyphs) => ContentElement::Text(TextSpan {
                        x: f64::from(glyphs.x),
                        y: f64::from(glyphs.y),
                        text: glyphs.text,
                        font_name: glyphs.face.key.clone(),
                        font_size: f64::from(glyphs.size),
                        stroke_width: if glyphs.fake_bold { f64::from(glyphs.size) * 0.03 } else { 0.0 },
                        skew: if glyphs.fake_italic { 0.2 } else { 0.0 },
                        fill_color: glyphs.color,
                        character_spacing: 0.0,
                        horizontal_scale: 100.0,
                        stroke_color: None,
                    }),
                })
                .collect(),
        })
        .collect();

    PdfDocument {
        version: PdfVersion { major: 1, minor: 7 },
        info: DocumentInfo::default(),
        pages,
        outline: Vec::new(),
        embedded_fonts,
    }
}

/// The pages of a presentation, one per slide, and what on them is not drawn
/// as PowerPoint draws it.
pub fn pages(pres: &Presentation) -> (PdfDocument, Vec<UnsupportedContent>) {
    let faces = Faces::default();
    let mut census = Census::default();
    let mut pages = Vec::with_capacity(pres.slides.len());
    for (number, slide) in pres.slides.iter().enumerate() {
        let mut draw = SlideDraw { pres, faces: &faces, census: &mut census, slide: number, items: Vec::new() };
        draw.background(slide);
        for shape in &slide.shapes {
            draw.shape(shape);
        }
        pages.push((f64::from(pres.slide_width), f64::from(pres.slide_height), draw.items));
    }
    (finish(pages, &faces), census.into_list())
}

#[cfg(test)]
mod tests {
    use super::*;

    fn fixture(name: &str) -> Presentation {
        let path = format!("{}/../../tests/fixtures/{name}", env!("CARGO_MANIFEST_DIR"));
        oxislides_core::parse_pptx(&std::fs::read(path).unwrap()).unwrap()
    }

    fn page_text(page: &Page) -> String {
        page.contents.iter().filter_map(|e| e.text_span()).map(|span| span.text.as_str()).collect::<Vec<_>>().join(" ")
    }

    #[test]
    fn each_slide_becomes_a_page_of_its_size() {
        let pres = fixture("multi_slide.pptx");
        let (doc, _) = pages(&pres);
        assert_eq!(doc.pages.len(), 3);
        for page in &doc.pages {
            assert!((page.width - f64::from(pres.slide_width)).abs() < 0.01);
            assert!((page.height - f64::from(pres.slide_height)).abs() < 0.01);
        }
    }

    #[test]
    fn a_slides_words_are_on_its_page_inside_the_slide() {
        let pres = fixture("basic_test.pptx");
        let (doc, _) = pages(&pres);
        let page = &doc.pages[0];
        let text = page_text(page);
        for word in ["Welcome", "to", "Oxi"] {
            assert!(text.contains(word), "{word:?} missing from {text:?}");
        }
        for span in page.contents.iter().filter_map(|e| e.text_span()) {
            assert!(span.x >= 0.0 && span.x <= page.width, "x {} off the slide", span.x);
            assert!(span.y >= 0.0 && span.y <= page.height, "y {} off the slide", span.y);
            assert!(doc.embedded_fonts.contains_key(&span.font_name), "{} not embedded", span.font_name);
        }
    }

    #[test]
    fn a_word_starts_a_new_span_so_faces_cannot_drift_the_line() {
        let pres = fixture("basic_test.pptx");
        let (doc, _) = pages(&pres);
        assert!(doc.pages[0].contents.iter().filter_map(|e| e.text_span()).all(|span| !span.text.contains(' ')));
    }

    #[test]
    fn numbering_follows_its_kind() {
        assert_eq!(auto_number("arabicPeriod", 3), "3.");
        assert_eq!(auto_number("alphaLcParenR", 28), "ab)");
        assert_eq!(auto_number("romanUcPeriod", 14), "XIV.");
        assert_eq!(auto_number("arabicParenBoth", 2), "(2)");
    }

    #[test]
    fn a_turned_box_turns_about_its_centre() {
        let at = Placement { x: 10.0, y: 20.0, w: 100.0, h: 40.0, flip_h: false, flip_v: false, sin: 1.0, cos: 0.0 };
        // A quarter turn clockwise carries the top-left corner to the
        // top-right of the turned box.
        let (x, y) = at.point(0.0, 0.0);
        assert!((x - 80.0).abs() < 1e-4 && (y + 10.0).abs() < 1e-4, "({x}, {y})");
    }

    #[test]
    fn a_translucent_fill_is_the_colour_it_shows_on_white() {
        let Color::Rgb(r, g, b) = over_white(Color::Rgb(0.0, 0.0, 1.0), 0.25) else { panic!() };
        assert!((r - 0.75).abs() < 1e-9 && (g - 0.75).abs() < 1e-9 && (b - 1.0).abs() < 1e-9);
    }

    fn end(kind: &str) -> oxislides_core::ir::LineEnd {
        oxislides_core::ir::LineEnd { kind: kind.to_string(), w: "med".to_string(), len: "med".to_string() }
    }

    #[test]
    fn a_triangle_end_has_its_tip_on_the_end_and_points_away_from_the_line() {
        // A 1pt line running right, ending at (100, 50): med x med on the
        // 2pt floor is 6pt along and 6pt across.
        let Some(Some(ops)) = line_end(&end("triangle"), (100.0, 50.0), (0.0, 50.0), 1.0) else { panic!() };
        let PathOp::MoveTo(x, y) = ops[0] else { panic!() };
        assert!((x - 100.0).abs() < 1e-9 && (y - 50.0).abs() < 1e-9);
        let PathOp::LineTo(bx, by) = ops[1] else { panic!() };
        assert!((bx - 94.0).abs() < 1e-9 && ((by - 50.0).abs() - 3.0).abs() < 1e-9, "({bx}, {by})");
    }

    #[test]
    fn an_open_arrow_end_is_not_drawn_but_none_is_quietly_nothing() {
        assert!(line_end(&end("arrow"), (1.0, 0.0), (0.0, 0.0), 1.0).is_none());
        assert!(matches!(line_end(&end("none"), (1.0, 0.0), (0.0, 0.0), 1.0), Some(None)));
    }

    #[test]
    fn an_elbow_connector_turns_at_its_adjustment() {
        let mut shape = fixture("basic_test.pptx").slides[0].shapes[0].clone();
        shape.shape_type = Some("bentConnector3".to_string());
        shape.custom_geometry = None;
        shape.width = 100.0;
        shape.height = 40.0;
        shape.adjustments.insert("adj1".to_string(), 25_000.0);
        let parts = outlines(&shape).unwrap();
        assert!(parts[0].0.open);
        let at = Placement { x: 0.0, y: 0.0, w: 100.0, h: 40.0, flip_h: false, flip_v: false, sin: 0.0, cos: 1.0 };
        assert_eq!(parts[0].0.points(&at), vec![(0.0, 0.0), (25.0, 0.0), (25.0, 40.0), (100.0, 40.0)]);
    }

    fn callout(kind: &str, adj1: f32, adj2: f32) -> Outline {
        let mut shape = Shape { width: 200.0, height: 100.0, shape_type: Some(kind.to_string()), ..Default::default() };
        shape.adjustments.insert("adj1".to_string(), adj1);
        shape.adjustments.insert("adj2".to_string(), adj2);
        outlines(&shape).unwrap().remove(0).0
    }

    fn has_point(outline: &Outline, (x, y): (f32, f32)) -> bool {
        outline.ops.iter().any(|op| matches!(*op, GeomCmd::LineTo(px, py) if (px - x).abs() < 1e-3 && (py - y).abs() < 1e-3))
    }

    #[test]
    fn a_callout_reaches_its_point() {
        // adj1 / adj2 are fractions of the box from its centre.
        for kind in ["wedgeEllipseCallout", "wedgeRectCallout", "wedgeRoundRectCallout"] {
            let outline = callout(kind, -20_833.0, 62_500.0);
            assert!(has_point(&outline, (100.0 - 41.666, 112.5)), "{kind}");
            assert!(!outline.open);
        }
    }

    #[test]
    fn an_oval_callout_opens_toward_its_point() {
        // Point straight below: the arc starts and ends 11 degrees either side
        // of the bottom, so both ends sit low on the ellipse, either side of
        // the centre line.
        let outline = callout("wedgeEllipseCallout", 0.0, 80_000.0);
        let GeomCmd::MoveTo(x0, y0) = outline.ops[0] else { panic!() };
        let end = outline.ops.iter().rev().find_map(|op| match *op {
            GeomCmd::CubicTo(.., x, y) => Some((x, y)),
            _ => None,
        });
        let (x1, y1) = end.unwrap();
        assert!(y0 > 95.0 && y1 > 95.0, "{y0} {y1}");
        assert!((x0 - 100.0) * (x1 - 100.0) < 0.0, "{x0} {x1}");
    }

    #[test]
    fn a_square_callout_takes_its_wedge_from_the_side_the_point_is_past() {
        // Far to the right and only a little down: the wedge leaves the right
        // side, between 7/12 and 10/12 of its height.
        let outline = callout("wedgeRectCallout", 80_000.0, 10_000.0);
        assert!(has_point(&outline, (200.0, 100.0 * 7.0 / 12.0)));
        assert!(has_point(&outline, (200.0, 100.0 * 10.0 / 12.0)));
        assert!(!has_point(&outline, (200.0 * 7.0 / 12.0, 100.0)));
    }

    #[test]
    fn a_folded_corner_draws_the_fold_darker() {
        let mut shape = fixture("basic_test.pptx").slides[0].shapes[0].clone();
        shape.shape_type = Some("foldedCorner".to_string());
        shape.custom_geometry = None;
        let parts = outlines(&shape).unwrap();
        assert_eq!(parts.len(), 2);
        assert!(parts[1].1 == PartFill::DarkenLess);
    }

    #[test]
    fn a_gradient_fill_is_reported() {
        let pres = fixture("gradalpha_test.pptx");
        let (_, unsupported) = pages(&pres);
        assert!(unsupported.iter().any(|u| u.kind.contains("gradient")), "{unsupported:?}");
    }
}
// Text placement shared with the spreadsheet PDF path. Font lookup and
// subsetting stay in one place so Japanese fallback fonts are embedded too.
#[derive(Default)]
pub(super) struct TextPages {
    faces: Faces,
    pages: Vec<(f64, f64, Vec<Item>)>,
}

impl TextPages {
    pub fn page(&mut self, width: f64, height: f64) {
        self.pages.push((width, height, Vec::new()));
    }
    pub fn element(&mut self, element: ContentElement) {
        self.pages.last_mut().unwrap().2.push(Item::Element(element));
    }
    pub fn width(&self, text: &str, family: &str, size: f64, bold: bool, italic: bool) -> f64 {
        text.chars().map(|ch| self.faces.face_for(family, bold, italic, ch)
            .map_or(0.5, |(_, advance)| f64::from(advance)) * size).sum()
    }
    pub fn text(&mut self, text: &str, family: &str, size: f64, bold: bool, italic: bool, x: f64, y: f64, color: Color) {
        let mut pen = x;
        let mut open: Option<Glyphs> = None;
        for ch in text.chars() {
            let Some((face, advance)) = self.faces.face_for(family, bold, italic, ch) else { continue };
            if let Some(run) = open.as_mut().filter(|run| run.face.key == face.key) {
                run.text.push(ch);
            } else {
                if let Some(run) = open.take() { self.pages.last_mut().unwrap().2.push(Item::Glyphs(run)); }
                open = Some(Glyphs { face, x: pen as f32, y: y as f32, text: ch.to_string(),
                    size: size as f32, color, fake_bold: false, fake_italic: false });
            }
            pen += f64::from(advance) * size;
        }
        if let Some(run) = open { self.pages.last_mut().unwrap().2.push(Item::Glyphs(run)); }
    }
    pub fn finish(self) -> PdfDocument { finish(self.pages, &self.faces) }
}
