// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! A page of the PDF IR drawn to pixels, for thumbnails.
//!
//! It reads the same `PdfDocument` that `oxipdf_core::write_pdf` writes, and
//! draws it the way a PDF viewer reads what the writer put down: glyphs from
//! the embedded font programs, advanced by the widths the writer puts in /W
//! (default 1000), with Tc, Tz and the synthetic-italic Tm shear; paths filled
//! non-zero and stroked; images scaled into their boxes; clips intersected and
//! scoped by q/Q. So a thumbnail shows exactly what the PDF shows, without
//! writing the PDF and reading it back.

use std::collections::HashMap;

use oxipdf_core::ir::{
    Color, ColorSpace, ContentElement, EmbeddedFont, FontFormat, ImageData, LineCap, LineJoin, Page, PathData, PathOp,
    PdfDocument, TextSpan,
};
use rustybuzz::ttf_parser;
use tiny_skia::{
    FillRule, Mask, Paint, PathBuilder, Pixmap, PixmapPaint, Stroke, Transform,
};

/// The size to draw a page at, in pixels. With both bounds the page fits
/// inside the box, aspect kept; with one, the other follows the page.
#[derive(Clone, Copy, Debug)]
pub struct Size {
    pub width: Option<u32>,
    pub height: Option<u32>,
}

impl Size {
    fn scale(&self, page: &Page) -> f32 {
        let (w, h) = (page.width.max(1.0) as f32, page.height.max(1.0) as f32);
        let by_w = self.width.map(|px| px as f32 / w);
        let by_h = self.height.map(|px| px as f32 / h);
        match (by_w, by_h) {
            (Some(a), Some(b)) => a.min(b),
            (Some(a), None) | (None, Some(a)) => a,
            // 96 dpi, as a screen shows a page at 100%.
            (None, None) => 96.0 / 72.0,
        }
    }
}

/// Glyph outlines of one font: a whole TrueType/OpenType file, or the bare
/// CFF program the writer embeds as FontFile3.
enum Outlines<'a> {
    Face(Box<ttf_parser::Face<'a>>),
    Cff(Box<ttf_parser::cff::Table<'a>>),
}

/// A font as the page draws it: outlines, the font matrix that takes its
/// units to one em, and the advance of each glyph in 1/1000 em.
struct Font<'a> {
    outlines: Outlines<'a>,
    /// sx, ky, kx, sy, tx, ty: font units to text space, before the size.
    matrix: [f32; 6],
    /// Glyph → advance in 1/1000 em; a glyph not listed advances `default_width`.
    widths: Widths<'a>,
    /// Character → glyph, for spans that carry text without glyph indices.
    cmap: Cmap<'a>,
}

enum Widths<'a> {
    /// The writer's /W: what the PDF advances by.
    Embedded(&'a HashMap<u16, u16>),
    /// A font the writer left unembedded, drawn with a stand-in face: its own
    /// advances.
    Face,
}

enum Cmap<'a> {
    Embedded(&'a HashMap<u32, u16>),
    Face,
}

impl<'a> Font<'a> {
    fn embedded(font: &'a EmbeddedFont) -> Option<Self> {
        let outlines = match font.format {
            FontFormat::OpenTypeCff => Outlines::Cff(Box::new(ttf_parser::cff::Table::parse(&font.data)?)),
            FontFormat::TrueType => Outlines::Face(Box::new(ttf_parser::Face::parse(&font.data, 0).ok()?)),
        };
        let matrix = matrix_of(&outlines);
        Some(Self { outlines, matrix, widths: Widths::Embedded(&font.cid_widths), cmap: Cmap::Embedded(&font.unicode_to_gid) })
    }

    fn stand_in(data: &'a [u8]) -> Option<Self> {
        let outlines = Outlines::Face(Box::new(ttf_parser::Face::parse(data, 0).ok()?));
        let matrix = matrix_of(&outlines);
        Some(Self { outlines, matrix, widths: Widths::Face, cmap: Cmap::Face })
    }

    /// The glyphs a span draws, as the writer encodes them.
    fn glyphs(&self, span: &TextSpan, selected: Option<&[u16]>) -> Vec<u16> {
        if let Some(indices) = selected {
            return indices.to_vec();
        }
        span.text
            .chars()
            .map(|ch| match &self.cmap {
                Cmap::Embedded(map) if map.is_empty() => ch as u16,
                Cmap::Embedded(map) => map.get(&(ch as u32)).copied().unwrap_or(0),
                Cmap::Face => match &self.outlines {
                    Outlines::Face(face) => face.glyph_index(ch).map_or(0, |g| g.0),
                    Outlines::Cff(_) => 0,
                },
            })
            .collect()
    }

    /// Advance of `gid` in 1/1000 em.
    fn width(&self, gid: u16) -> f32 {
        match (&self.widths, &self.outlines) {
            (Widths::Embedded(map), _) => map.get(&gid).copied().unwrap_or(1000) as f32,
            (Widths::Face, Outlines::Face(face)) => {
                let upem = face.units_per_em().max(1) as f32;
                face.glyph_hor_advance(ttf_parser::GlyphId(gid)).unwrap_or(0) as f32 * 1000.0 / upem
            }
            (Widths::Face, Outlines::Cff(_)) => 1000.0,
        }
    }

    fn outline(&self, gid: u16) -> Option<tiny_skia::Path> {
        let mut sink = Sink(PathBuilder::new());
        let id = ttf_parser::GlyphId(gid);
        match &self.outlines {
            Outlines::Face(face) => face.outline_glyph(id, &mut sink)?,
            Outlines::Cff(cff) => cff.outline(id, &mut sink).ok()?,
        };
        sink.0.finish()
    }
}

fn matrix_of(outlines: &Outlines) -> [f32; 6] {
    match outlines {
        Outlines::Face(face) => {
            let k = 1.0 / face.units_per_em().max(1) as f32;
            [k, 0.0, 0.0, k, 0.0, 0.0]
        }
        Outlines::Cff(cff) => {
            let m = cff.matrix();
            [m.sx, m.ky, m.kx, m.sy, m.tx, m.ty]
        }
    }
}

struct Sink(PathBuilder);

impl ttf_parser::OutlineBuilder for Sink {
    fn move_to(&mut self, x: f32, y: f32) {
        self.0.move_to(x, y);
    }
    fn line_to(&mut self, x: f32, y: f32) {
        self.0.line_to(x, y);
    }
    fn quad_to(&mut self, x1: f32, y1: f32, x: f32, y: f32) {
        self.0.quad_to(x1, y1, x, y);
    }
    fn curve_to(&mut self, x1: f32, y1: f32, x2: f32, y2: f32, x: f32, y: f32) {
        self.0.cubic_to(x1, y1, x2, y2, x, y);
    }
    fn close(&mut self) {
        self.0.close();
    }
}

/// Draw page `index` (0-based) of `doc`. `stand_in` is a TrueType face for
/// text in fonts the writer did not embed (the PDF names a standard font a
/// viewer substitutes; this is ours). None when the page does not exist or
/// the size is degenerate.
pub fn render_page(doc: &PdfDocument, index: usize, size: Size, stand_in: Option<&[u8]>) -> Option<Pixmap> {
    let page = doc.pages.get(index)?;
    let scale = size.scale(page);
    let px_w = (page.width as f32 * scale).round().max(1.0) as u32;
    let px_h = (page.height as f32 * scale).round().max(1.0) as u32;
    let mut pixmap = Pixmap::new(px_w, px_h)?;
    pixmap.fill(tiny_skia::Color::WHITE);

    let mut fonts: HashMap<&str, Option<Font>> = HashMap::new();
    let stand_in_font = stand_in.and_then(Font::stand_in);
    let base = Transform::from_scale(scale, scale);
    // q/Q scope the clip and nothing else the IR sets: every element carries
    // its own colours and widths.
    let mut clip: Option<Mask> = None;
    let mut saved: Vec<Option<Mask>> = Vec::new();

    for element in &page.contents {
        match element {
            ContentElement::SaveState => saved.push(clip.clone()),
            ContentElement::RestoreState => clip = saved.pop().unwrap_or(None),
            ContentElement::ClipPath(c) => {
                let Some(path) = build_path(&c.operations) else {
                    // An empty clip shows nothing.
                    clip = Mask::new(px_w, px_h);
                    continue;
                };
                let rule = if c.even_odd { FillRule::EvenOdd } else { FillRule::Winding };
                match clip.as_mut() {
                    Some(mask) => mask.intersect_path(&path, rule, true, base),
                    None => {
                        let mut mask = Mask::new(px_w, px_h)?;
                        mask.fill_path(&path, rule, true, base);
                        clip = Some(mask);
                    }
                }
            }
            ContentElement::Path(p) => draw_path(&mut pixmap, p, base, clip.as_ref()),
            ContentElement::Image(img) => draw_image(&mut pixmap, img, base, clip.as_ref()),
            ContentElement::Text(_) | ContentElement::GlyphRun(_) => {
                let span = element.text_span().expect("text element");
                let font = fonts
                    .entry(span.font_name.as_str())
                    .or_insert_with(|| doc.embedded_fonts.get(&span.font_name).and_then(Font::embedded));
                let Some(font) = font.as_ref().or(stand_in_font.as_ref()) else { continue };
                draw_text(&mut pixmap, span, element.glyph_indices(), font, base, clip.as_ref());
            }
        }
    }
    Some(pixmap)
}

fn draw_text(pixmap: &mut Pixmap, span: &TextSpan, selected: Option<&[u16]>, font: &Font, base: Transform, clip: Option<&Mask>) {
    let fs = span.font_size as f32;
    let th = if span.horizontal_scale > 0.0 { span.horizontal_scale as f32 / 100.0 } else { 1.0 };
    let skew = span.skew as f32;
    let [a, b, c, d, e, f] = font.matrix;
    let fill = paint(&span.fill_color);
    let stroke = (span.stroke_width > 0.0).then(|| {
        (paint(span.stroke_color.as_ref().unwrap_or(&span.fill_color)), Stroke { width: span.stroke_width as f32, ..Stroke::default() })
    });
    let mut pen = 0.0f32;
    for gid in font.glyphs(span, selected) {
        if let Some(path) = font.outline(gid) {
            // Font units → text space (matrix × size, Tz on x) → page (the Tm
            // shear, then y flipped: the IR's y runs down from the top).
            let glyph = Transform::from_row(
                fs * (th * a + skew * b),
                -fs * b,
                fs * (th * c + skew * d),
                -fs * d,
                span.x as f32 + pen + fs * (th * e + skew * f),
                span.y as f32 - fs * f,
            );
            let t = base.pre_concat(glyph);
            pixmap.fill_path(&path, &fill, FillRule::Winding, t, clip);
            if let Some((stroke_paint, stroke)) = &stroke {
                // Line width is in page units: stroke the outline mapped to
                // the page, not in font units.
                if let Some(page_path) = path.clone().transform(glyph) {
                    pixmap.stroke_path(&page_path, stroke_paint, stroke, base, clip);
                }
            }
        }
        pen += (font.width(gid) / 1000.0 * fs) * th + span.character_spacing as f32;
    }
}

fn draw_path(pixmap: &mut Pixmap, p: &PathData, base: Transform, clip: Option<&Mask>) {
    let Some(path) = build_path(&p.operations) else { return };
    if let Some(fill) = &p.fill {
        pixmap.fill_path(&path, &paint(fill), FillRule::Winding, base, clip);
    }
    if let Some(s) = &p.stroke {
        let stroke = Stroke {
            // Viewers draw no line thinner than a device pixel (PDF's 0 means
            // exactly that): a 0.5pt table rule at thumbnail size would
            // otherwise fade to a grey wash.
            width: (s.width as f32).max(1.0 / base.sx),
            line_cap: match s.line_cap {
                LineCap::Butt => tiny_skia::LineCap::Butt,
                LineCap::Round => tiny_skia::LineCap::Round,
                LineCap::Square => tiny_skia::LineCap::Square,
            },
            line_join: match s.line_join {
                LineJoin::Miter => tiny_skia::LineJoin::Miter,
                LineJoin::Round => tiny_skia::LineJoin::Round,
                LineJoin::Bevel => tiny_skia::LineJoin::Bevel,
            },
            ..Stroke::default()
        };
        pixmap.stroke_path(&path, &paint(&s.color), &stroke, base, clip);
    }
}

fn draw_image(pixmap: &mut Pixmap, img: &ImageData, base: Transform, clip: Option<&Mask>) {
    let (w, h) = (img.pixel_width, img.pixel_height);
    if w == 0 || h == 0 || img.width <= 0.0 || img.height <= 0.0 || img.bits_per_component != 8 {
        return;
    }
    let channels = match img.color_space {
        ColorSpace::DeviceGray => 1,
        ColorSpace::DeviceRgb => 3,
        ColorSpace::DeviceCmyk => 4,
    };
    let pixels = (w as usize) * (h as usize);
    if img.data.len() < pixels * channels {
        return;
    }
    let Some(mut source) = Pixmap::new(w, h) else { return };
    for (i, out) in source.data_mut().chunks_exact_mut(4).enumerate() {
        let px = &img.data[i * channels..(i + 1) * channels];
        let (r, g, b) = match img.color_space {
            ColorSpace::DeviceGray => (px[0], px[0], px[0]),
            ColorSpace::DeviceRgb => (px[0], px[1], px[2]),
            ColorSpace::DeviceCmyk => {
                let k = 255 - px[3] as u16;
                let ch = |v: u8| ((255 - v as u16) * k / 255) as u8;
                (ch(px[0]), ch(px[1]), ch(px[2]))
            }
        };
        let alpha = img.alpha.as_ref().and_then(|a| a.get(i).copied()).unwrap_or(255);
        // Premultiplied, as tiny-skia stores pixels.
        let pm = |v: u8| ((v as u16 * alpha as u16 + 127) / 255) as u8;
        out.copy_from_slice(&[pm(r), pm(g), pm(b), alpha]);
    }
    let place = Transform::from_row(img.width as f32 / w as f32, 0.0, 0.0, img.height as f32 / h as f32, img.x as f32, img.y as f32);
    let quality = tiny_skia::FilterQuality::Bicubic;
    pixmap.draw_pixmap(0, 0, source.as_ref(), &PixmapPaint { quality, ..PixmapPaint::default() }, base.pre_concat(place), clip);
}

fn build_path(ops: &[PathOp]) -> Option<tiny_skia::Path> {
    let mut pb = PathBuilder::new();
    for op in ops {
        match *op {
            PathOp::MoveTo(x, y) => pb.move_to(x as f32, y as f32),
            PathOp::LineTo(x, y) => pb.line_to(x as f32, y as f32),
            PathOp::CurveTo(x1, y1, x2, y2, x, y) => {
                pb.cubic_to(x1 as f32, y1 as f32, x2 as f32, y2 as f32, x as f32, y as f32)
            }
            PathOp::ClosePath => pb.close(),
        }
    }
    pb.finish()
}

fn paint(color: &Color) -> Paint<'static> {
    let (r, g, b) = match *color {
        Color::Gray(v) => (v, v, v),
        Color::Rgb(r, g, b) => (r, g, b),
        Color::Cmyk(c, m, y, k) => ((1.0 - c) * (1.0 - k), (1.0 - m) * (1.0 - k), (1.0 - y) * (1.0 - k)),
    };
    let unit = |v: f64| v.clamp(0.0, 1.0) as f32;
    let mut paint = Paint::default();
    paint.set_color(tiny_skia::Color::from_rgba(unit(r), unit(g), unit(b), 1.0).expect("clamped"));
    paint.anti_alias = true;
    paint
}

/// The pixmap as PNG bytes (opaque: pages are drawn on white).
pub fn encode_png(pixmap: &Pixmap) -> Result<Vec<u8>, String> {
    let mut rgb = Vec::with_capacity((pixmap.width() * pixmap.height() * 3) as usize);
    for px in pixmap.data().chunks_exact(4) {
        rgb.extend_from_slice(&px[..3]);
    }
    let mut out = std::io::Cursor::new(Vec::new());
    image::write_buffer_with_format(
        &mut out,
        &rgb,
        pixmap.width(),
        pixmap.height(),
        image::ExtendedColorType::Rgb8,
        image::ImageFormat::Png,
    )
    .map_err(|e| e.to_string())?;
    Ok(out.into_inner())
}

#[cfg(test)]
mod tests {
    use super::*;
    use oxipdf_core::ir::{DocumentInfo, PdfVersion, Rectangle, StrokeStyle};

    fn page(contents: Vec<ContentElement>) -> PdfDocument {
        PdfDocument {
            version: PdfVersion::new(1, 7),
            info: DocumentInfo::default(),
            pages: vec![Page {
                width: 100.0,
                height: 200.0,
                media_box: Rectangle { llx: 0.0, lly: 0.0, urx: 100.0, ury: 200.0 },
                crop_box: None,
                contents,
                rotation: 0,
            }],
            outline: vec![],
            embedded_fonts: HashMap::new(),
        }
    }

    fn rect(x: f64, y: f64, w: f64, h: f64) -> Vec<PathOp> {
        vec![
            PathOp::MoveTo(x, y),
            PathOp::LineTo(x + w, y),
            PathOp::LineTo(x + w, y + h),
            PathOp::LineTo(x, y + h),
            PathOp::ClosePath,
        ]
    }

    fn pixel(p: &Pixmap, x: u32, y: u32) -> [u8; 3] {
        let c = p.pixel(x, y).unwrap();
        [c.red(), c.green(), c.blue()]
    }

    #[test]
    fn size_fits_the_box_and_keeps_the_aspect() {
        let doc = page(vec![]);
        let fit = |w, h| {
            let p = render_page(&doc, 0, Size { width: w, height: h }, None).unwrap();
            (p.width(), p.height())
        };
        assert_eq!(fit(Some(50), None), (50, 100));
        assert_eq!(fit(None, Some(50)), (25, 50));
        assert_eq!(fit(Some(50), Some(50)), (25, 50));
        assert!(render_page(&doc, 1, Size { width: Some(50), height: None }, None).is_none());
    }

    #[test]
    fn a_filled_path_lands_where_the_ir_puts_it_y_down() {
        let doc = page(vec![ContentElement::Path(PathData {
            operations: rect(0.0, 0.0, 50.0, 20.0),
            stroke: None,
            fill: Some(Color::Rgb(1.0, 0.0, 0.0)),
        })]);
        let p = render_page(&doc, 0, Size { width: Some(100), height: None }, None).unwrap();
        assert_eq!(pixel(&p, 10, 10), [255, 0, 0]);
        assert_eq!(pixel(&p, 10, 190), [255, 255, 255]);
        assert_eq!(pixel(&p, 75, 10), [255, 255, 255]);
    }

    #[test]
    fn a_clip_holds_until_its_restore() {
        let fill = |c| {
            ContentElement::Path(PathData { operations: rect(0.0, 0.0, 100.0, 200.0), stroke: None, fill: Some(c) })
        };
        let doc = page(vec![
            ContentElement::SaveState,
            ContentElement::ClipPath(oxipdf_core::ir::ClipPathData { operations: rect(0.0, 0.0, 10.0, 10.0), even_odd: false }),
            fill(Color::Gray(0.0)),
            ContentElement::RestoreState,
            ContentElement::Path(PathData {
                operations: rect(0.0, 100.0, 100.0, 100.0),
                stroke: Some(StrokeStyle { color: Color::Gray(0.0), width: 1.0, line_cap: LineCap::Butt, line_join: LineJoin::Miter }),
                fill: Some(Color::Rgb(0.0, 0.0, 1.0)),
            }),
        ]);
        let p = render_page(&doc, 0, Size { width: Some(100), height: None }, None).unwrap();
        assert_eq!(pixel(&p, 5, 5), [0, 0, 0], "inside the clip");
        assert_eq!(pixel(&p, 50, 50), [255, 255, 255], "outside the clip");
        assert_eq!(pixel(&p, 50, 150), [0, 0, 255], "after the restore");
    }

    #[test]
    fn an_image_fills_its_box() {
        let doc = page(vec![ContentElement::Image(ImageData {
            x: 50.0,
            y: 100.0,
            width: 50.0,
            height: 100.0,
            data: vec![0, 255, 0, 0, 255, 0, 0, 255, 0, 0, 255, 0],
            color_space: ColorSpace::DeviceRgb,
            bits_per_component: 8,
            pixel_width: 2,
            pixel_height: 2,
            alpha: None,
        })]);
        let p = render_page(&doc, 0, Size { width: Some(100), height: None }, None).unwrap();
        assert_eq!(pixel(&p, 75, 150), [0, 255, 0]);
        assert_eq!(pixel(&p, 25, 50), [255, 255, 255]);
    }
}
