// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! Laid-out Word pages as the `PdfDocument` the PDF writer writes, drawn
//! with the faces Oxidocs ships (Carlito, Liberation, Caladea, OxiGothic /
//! OxiMincho and their proportional companions) or programs registered at
//! run time ([`crate::fonts`]) -- no font file is read from disk. The browser
//! build's PDF export and the native viewers both draw from here, so a page
//! looks the same in a tab, in the PDF and on a tablet.

use std::sync::OnceLock;

use oxidocs_fonts::{OXI_GOTHIC, OXI_GOTHIC_P, OXI_MINCHO, OXI_MINCHO_P};

/// Which bundled family stands in for a requested one. Anything not listed
/// keeps its own name and goes out unembedded, exactly as before.
pub fn latin_stand_in(base_font: &str) -> Option<&'static str> {
    Some(match base_font {
        "Calibri" => "Carlito",
        "Arial" | "Helvetica" | "Segoe UI" | "Tahoma" | "Verdana" => "LiberationSans",
        "Times New Roman" | "Times" | "Georgia" => "LiberationSerif",
        "Courier New" | "Courier" | "Consolas" => "LiberationMono",
        "Cambria" | "Century" => "Caladea",
        _ => return None,
    })
}

/// Parse one bundled face on first use and keep it. Parsing is cheap next to
/// the PDF write, but a document with many runs asks for the same face over
/// and over.
pub fn latin_face(key: &str) -> Option<&'static oxipdf_core::ir::EmbeddedFont> {
    static CACHE: OnceLock<
        std::collections::HashMap<&'static str, oxipdf_core::ir::EmbeddedFont>,
    > = OnceLock::new();
    let cache = CACHE.get_or_init(|| {
        oxidocs_fonts::latin_faces()
            .into_iter()
            .map(|(name, data)| (name, oxipdf_core::font_util::embedded_font_from_ttf(data)))
            .collect()
    });
    cache.get(key)
}

pub fn get_cjk_fonts() -> &'static (oxipdf_core::ir::EmbeddedFont, oxipdf_core::ir::EmbeddedFont) {
    static FONTS: OnceLock<(oxipdf_core::ir::EmbeddedFont, oxipdf_core::ir::EmbeddedFont)> = OnceLock::new();
    FONTS.get_or_init(|| {
        let gothic = oxipdf_core::font_util::embedded_font_from_ttf(OXI_GOTHIC);
        let mincho = oxipdf_core::font_util::embedded_font_from_ttf(OXI_MINCHO);
        (gothic, mincho)
    })
}

/// The proportional companions of the two stand-ins, keyed as the PDF export
/// names them.
pub fn get_cjk_proportional_fonts() -> &'static [(&'static str, oxipdf_core::ir::EmbeddedFont); 2] {
    static FONTS: OnceLock<[(&'static str, oxipdf_core::ir::EmbeddedFont); 2]> = OnceLock::new();
    FONTS.get_or_init(|| [
        ("OxiCJK-GothicP", oxipdf_core::font_util::embedded_font_from_ttf(OXI_GOTHIC_P)),
        ("OxiCJK-MinchoP", oxipdf_core::font_util::embedded_font_from_ttf(OXI_MINCHO_P)),
    ])
}

pub fn metrics_registry() -> &'static oxidocs_core::font::FontMetricsRegistry {
    static METRICS: OnceLock<oxidocs_core::font::FontMetricsRegistry> = OnceLock::new();
    METRICS.get_or_init(oxidocs_core::font::FontMetricsRegistry::load)
}

/// An element drawn with a stood-in East Asian face (`OxiCJK-Gothic` /
/// `OxiCJK-Mincho`) cut where what the family sets proportionally starts and
/// stops (`FontMetrics::draws_proportional`): (text, face key, x, horizontal
/// scale) per piece. The stand-ins set ASCII at half an em and kana at a full
/// em, as ＭＳ ゴシック does, so 游ゴシック or メイリオ came out in a typewriter
/// face with a gap after each word, and ＭＳ Ｐゴシック's half-em 「（」「・」 ran
/// into the next character; those are drawn from OxiGothicP / OxiMinchoP
/// instead, placed at the layout's measured advances and fitted to the width
/// measured for them -- Latin either way, kana and punctuation only narrowed
/// (see the CLI's `PROPORTIONAL`). None when nothing needs cutting.
pub fn proportional_pieces(
    text: &str, family: &str, face_key: &str, x: f64, width: f64, fs: f64, spacing: f64,
) -> Option<Vec<(String, String, f64, f64)>> {
    let companion_key = match face_key {
        "OxiCJK-Gothic" => "OxiCJK-GothicP",
        "OxiCJK-Mincho" => "OxiCJK-MinchoP",
        _ => return None,
    };
    let companion = &get_cjk_proportional_fonts().iter().find(|(key, _)| *key == companion_key)?.1;
    let metrics = metrics_registry().get(family);
    let proportional = |c: char| oxidocs_fonts::in_cjk_proportional_stand_in(c)
        && metrics.draws_proportional(c)
        && companion.unicode_to_gid.contains_key(&(c as u32));
    if !text.chars().any(proportional) {
        return None;
    }
    // A kana or punctuation mark from the companion is a piece of its own,
    // set in the cell the layout measured for it.
    let mut pieces: Vec<(String, bool, f64)> = Vec::new();
    for c in text.chars() {
        let is_proportional = proportional(c);
        let advance = metrics.char_width_em(c) as f64 * fs + spacing;
        match pieces.last_mut() {
            Some((piece, kind, w)) if *kind == is_proportional
                && !(is_proportional && (!c.is_ascii() || !piece.is_ascii())) =>
            {
                piece.push(c);
                *w += advance;
            }
            _ => pieces.push((c.to_string(), is_proportional, advance)),
        }
    }
    let total: f64 = pieces.iter().map(|(_, _, w)| w).sum();
    let scale = if total > 0.0 && width > 0.0 { width / total } else { 1.0 };
    let mut at = x;
    Some(pieces.into_iter().map(|(piece, is_proportional, w)| {
        let w = w * scale;
        let start = at;
        at += w;
        if !is_proportional {
            return (piece, face_key.to_string(), start, 100.0);
        }
        let natural: f64 = piece.chars()
            .filter_map(|c| companion.unicode_to_gid.get(&(c as u32)))
            .map(|g| companion.cid_widths.get(g).copied().unwrap_or(1000) as f64 / 1000.0 * fs)
            .sum();
        let target = w - piece.chars().count() as f64 * spacing;
        let (fit, shift) = proportional_fit(&piece, natural, target);
        (piece, companion_key.to_string(), start + shift, 100.0 * fit)
    }).collect())
}

/// How a piece from the proportional companion meets the `target` width the
/// layout measured for it, given its `natural` width: (scale factor, x shift).
/// As the CLI's `proportional_fit`: a Latin word is scaled to the width either
/// way within 0.6..=1.4; a kana or mark is only narrowed, to 0.6 at most, and
/// centred on its cell when still wider than it (ＭＳ Ｐゴシック's «ｉ» is a
/// quarter em, Noto's over half).
pub fn proportional_fit(text: &str, natural: f64, target: f64) -> (f64, f64) {
    if natural <= 0.0 || target <= 0.0 {
        return (1.0, 0.0);
    }
    let fit = target / natural;
    if text.is_ascii() {
        return (if (0.6..=1.4).contains(&fit) { fit } else { 1.0 }, 0.0);
    }
    let fit = fit.clamp(0.6, 1.0);
    (fit, ((target - natural * fit) / 2.0).min(0.0))
}

#[cfg(test)]
mod proportional_tests {
    use super::proportional_pieces;

    #[test]
    fn the_latin_of_a_proportional_family_is_cut_out() {
        let pieces = proportional_pieces("游Windows・", "游ゴシック", "OxiCJK-Gothic", 10.0, 60.0, 10.5, 0.0).unwrap();
        let keys: Vec<&str> = pieces.iter().map(|(_, key, _, _)| key.as_str()).collect();
        assert_eq!(keys, ["OxiCJK-Gothic", "OxiCJK-GothicP", "OxiCJK-Gothic"]);
        assert_eq!(pieces[1].0, "Windows");
        assert_eq!(pieces[0].2, 10.0);
        assert!(pieces[0].2 < pieces[1].2 && pieces[1].2 < pieces[2].2 && pieces[2].2 < 70.0);
        let mincho = proportional_pieces("Office", "游明朝", "OxiCJK-Mincho", 0.0, 30.0, 10.5, 0.0).unwrap();
        assert_eq!(mincho[0].1, "OxiCJK-MinchoP");
    }

    /// ＭＳ Ｐゴシック's narrow kana and punctuation come from the companion, its
    /// kanji and full-em kana (が, measured a full em) from the main face, each
    /// piece where the measured widths put it and the kana narrowed to them.
    #[test]
    fn the_kana_of_a_proportional_kana_family_are_cut_out() {
        let fs = 10.0;
        let m = super::metrics_registry().get("ＭＳ Ｐゴシック");
        let w = |s: &str| s.chars().map(|c| m.char_width_em(c) as f64 * fs).sum::<f64>();
        let pieces = proportional_pieces("（何ができる）", "ＭＳ Ｐゴシック", "OxiCJK-Gothic", 0.0, 0.0, fs, 0.0).unwrap();
        let texts: Vec<&str> = pieces.iter().map(|(t, _, _, _)| t.as_str()).collect();
        assert_eq!(texts, ["（", "何が", "で", "き", "る", "）"]);
        let keys: Vec<&str> = pieces.iter().map(|(_, k, _, _)| k.as_str()).collect();
        assert_eq!(keys, ["OxiCJK-GothicP", "OxiCJK-Gothic", "OxiCJK-GothicP", "OxiCJK-GothicP", "OxiCJK-GothicP", "OxiCJK-GothicP"]);
        assert!((pieces[1].2 - 5.0).abs() < 0.01, "何 starts after the half-em （: {}", pieces[1].2);
        assert!((pieces[2].2 - w("（何が")).abs() < 0.01);
        assert!((pieces[5].2 - w("（何ができる")).abs() < 0.01, "each kana in its own measured cell");
        assert_eq!(pieces[0].3, 100.0, "（ is half an em in both");
        assert!(pieces[2].3 < 100.0 && pieces[2].3 > 60.0, "the kana are narrowed: {}", pieces[2].3);
        // Latin stays one piece beside a mark.
        let latin = proportional_pieces("「Oxi」", "ＭＳ Ｐゴシック", "OxiCJK-Gothic", 0.0, 0.0, fs, 0.0).unwrap();
        let texts: Vec<&str> = latin.iter().map(|(t, _, _, _)| t.as_str()).collect();
        assert_eq!(texts, ["「", "Oxi", "」"]);
    }

    #[test]
    fn half_width_latin_and_other_faces_are_left_whole() {
        assert!(proportional_pieces("Windows", "ＭＳ ゴシック", "OxiCJK-Gothic", 0.0, 36.75, 10.5, 0.0).is_none());
        assert!(proportional_pieces("Windows", "ＭＳ 明朝", "OxiCJK-Mincho", 0.0, 36.75, 10.5, 0.0).is_none());
        assert!(proportional_pieces("あいう（）", "游ゴシック", "OxiCJK-Gothic", 0.0, 31.5, 10.5, 0.0).is_none());
        assert!(proportional_pieces("（あ）", "ＭＳ ゴシック", "OxiCJK-Gothic", 0.0, 31.5, 10.5, 0.0).is_none());
        assert!(proportional_pieces("Windows", "游ゴシック", "Carlito", 0.0, 40.0, 10.5, 0.0).is_none());
    }
}

// A baseline offset is optional in layout when a line needs no shared baseline.
// Canvas still needs an alphabetic origin. Resolve that origin from the same
// font tables and layout-origin frame used by the native horizontal renderer.
// Explicit line baselines win; vertical text keeps its existing placement.
pub fn canvas_baseline_offset(elem: &oxidocs_core::layout::LayoutElement) -> Option<f32> {
    match &elem.content {
        oxidocs_core::layout::LayoutContent::Text {
            font_size, font_family, is_vertical: false, ..
        } => {
            let registry = metrics_registry();
            Some(elem.baseline_offset.unwrap_or_else(|| {
                let metrics = registry.get(font_family.as_deref().unwrap_or("Calibri"));
                elem.text_y_off - 1.0 + metrics.baseline_ascent() * *font_size
            }))
        }
        _ => elem.baseline_offset,
    }
}

pub fn layout_to_pdf(
    layout: &oxidocs_core::layout::LayoutResult,
    doc: &oxidocs_core::Document,
    title: Option<String>,
) -> oxipdf_core::ir::PdfDocument {
    use oxipdf_core::ir::*;

    let title = title.or_else(|| doc.pages.first()
        .and_then(|p| p.blocks.first())
        .and_then(|b| match b {
            oxidocs_core::ir::Block::Paragraph(p) if p.style.heading_level == Some(1) => {
                Some(p.runs.iter().map(|r| r.text.as_str()).collect::<String>())
            }
            _ => None,
        }));

    let pages = layout.pages.iter().map(|lp| {
        let width = lp.width as f64;
        let height = lp.height as f64;
        let mut contents = Vec::new();

        for elem in &lp.elements {
            match &elem.content {
                oxidocs_core::layout::LayoutContent::Text {
                    text, font_size, font_family, bold, italic, color, character_spacing,
                    underline, underline_style, strikethrough, ..
                } => {
                    if text.is_empty() {
                        continue;
                    }
                    let base_font = font_family.as_deref().unwrap_or("Helvetica");
                    let registered = crate::fonts::has_font_program(base_font, *bold, *italic);
                    let font_name = if registered {
                        crate::fonts::resource_name(base_font, *bold, *italic)
                    } else {
                        resolve_wasm_font(base_font, *bold, *italic, text)
                    };

                    let fill_color = color.as_ref()
                        .and_then(|c| parse_hex_color(c))
                        .unwrap_or(Color::Gray(0.0));
                    let fs = *font_size as f64;
                    let baseline = (elem.y + canvas_baseline_offset(elem).unwrap_or(0.0)) as f64;

                    // The bundled East Asian stand-ins have no bold member, so bold
                    // East Asian text is drawn as Word draws a family without one:
                    // the regular face filled and stroked at fs/35 (the CLI's
                    // OXI_PDF_FAKE_BOLD; Word's PDF writes «2 Tr 0.51429 w» at 18pt).
                    let fake_bold = *bold && !registered && font_name.starts_with("OxiCJK-");

                    let pieces = if registered || elem.font_glyph.is_some() { None } else {
                        proportional_pieces(text, base_font, &font_name, elem.x as f64,
                            elem.width as f64, fs, *character_spacing as f64)
                    };
                    let split = pieces.is_some();
                    if let Some(pieces) = pieces {
                        for (piece, key, x, hscale) in pieces {
                            contents.push(ContentElement::Text(TextSpan {
                                x,
                                y: baseline,
                                text: piece,
                                font_name: key,
                                font_size: fs,
                                fill_color,
                                character_spacing: *character_spacing as f64,
                                stroke_width: if fake_bold { fs / 35.0 } else { 0.0 },
                                skew: 0.0, horizontal_scale: hscale, stroke_color: None,
                            }));
                        }
                    }
                    let span = TextSpan {
                        x: elem.x as f64,
                        y: baseline,
                        text: text.clone(),
                        font_name,
                        font_size: fs,
                        fill_color,
                        character_spacing: *character_spacing as f64,
                        stroke_width: if fake_bold { fs / 35.0 } else { 0.0 },
                        skew: 0.0, horizontal_scale: 100.0, stroke_color: None,
                    };
                    if !split {
                        contents.push(if let Some(glyph) = &elem.font_glyph {
                            ContentElement::GlyphRun(GlyphRun { span, glyph_indices: vec![glyph.index] })
                        } else {
                            ContentElement::Text(span)
                        });
                    }

                    // Underline and strikethrough across the element's width, at the
                    // positions and thicknesses the CLI's PDF export uses.
                    let (x0, x1) = (elem.x as f64, (elem.x + elem.width) as f64);
                    let rule = |y: f64, width: f64| ContentElement::Path(PathData {
                        operations: vec![PathOp::MoveTo(x0, y), PathOp::LineTo(x1, y)],
                        stroke: Some(StrokeStyle {
                            color: fill_color, width,
                            line_cap: LineCap::Butt, line_join: LineJoin::Miter,
                        }),
                        fill: None,
                    });
                    if *underline && x1 > x0 {
                        let w = 0.5_f64.max(fs * 0.05);
                        if underline_style.as_deref() == Some("double") {
                            contents.push(rule(baseline + fs * 0.10, w));
                            contents.push(rule(baseline + fs * 0.22, w));
                        } else {
                            contents.push(rule(baseline + fs * 0.15, w));
                        }
                    }
                    if *strikethrough && x1 > x0 {
                        contents.push(rule(elem.y as f64 + elem.height as f64 * 0.45, 0.5_f64.max(fs * 0.04)));
                    }
                }
                oxidocs_core::layout::LayoutContent::Image { data, .. } => {
                    if data.is_empty() {
                        continue;
                    }
                    contents.push(ContentElement::Image(ImageData {
                        x: elem.x as f64,
                        y: elem.y as f64,
                        width: elem.width as f64,
                        height: elem.height as f64,
                        data: data.clone(),
                        color_space: ColorSpace::DeviceRgb,
                        bits_per_component: 8,
                        pixel_width: elem.width as u32,
                        pixel_height: elem.height as u32,
                        alpha: None,
                    }));
                }
                // WordArt watermark: not drawn in the wasm PDF export yet
                oxidocs_core::layout::LayoutContent::WatermarkText { .. } => {}
                oxidocs_core::layout::LayoutContent::TableBorder { x1, y1, x2, y2, ref color, width, .. } => {
                    let stroke_color = color.as_ref()
                        .and_then(|c| parse_hex_color(c))
                        .unwrap_or(Color::Gray(0.0));
                    contents.push(ContentElement::Path(PathData {
                        operations: vec![
                            PathOp::MoveTo(*x1 as f64, *y1 as f64),
                            PathOp::LineTo(*x2 as f64, *y2 as f64),
                        ],
                        stroke: Some(StrokeStyle {
                            color: stroke_color,
                            width: *width as f64,
                            line_cap: LineCap::Butt,
                            line_join: LineJoin::Miter,
                        }),
                        fill: None,
                    }));
                }
                oxidocs_core::layout::LayoutContent::CellShading { ref color } => {
                    if let Some(fill_color) = parse_hex_color(color) {
                        contents.push(ContentElement::Path(PathData {
                            operations: vec![
                                PathOp::MoveTo(elem.x as f64, elem.y as f64),
                                PathOp::LineTo((elem.x + elem.width) as f64, elem.y as f64),
                                PathOp::LineTo((elem.x + elem.width) as f64, (elem.y + elem.height) as f64),
                                PathOp::LineTo(elem.x as f64, (elem.y + elem.height) as f64),
                                PathOp::ClosePath,
                            ],
                            stroke: None,
                            fill: Some(fill_color),
                        }));
                    }
                }
                oxidocs_core::layout::LayoutContent::BoxRect { ref fill, ref stroke_color, stroke_width, .. } => {
                    let fill_color = fill.as_deref().and_then(parse_hex_color);
                    let stroke = stroke_color.as_deref().and_then(parse_hex_color).map(|c| StrokeStyle {
                        color: c, width: *stroke_width as f64,
                        line_cap: LineCap::Butt, line_join: LineJoin::Miter,
                    });
                    contents.push(ContentElement::Path(PathData {
                        operations: vec![
                            PathOp::MoveTo(elem.x as f64, elem.y as f64),
                            PathOp::LineTo((elem.x + elem.width) as f64, elem.y as f64),
                            PathOp::LineTo((elem.x + elem.width) as f64, (elem.y + elem.height) as f64),
                            PathOp::LineTo(elem.x as f64, (elem.y + elem.height) as f64),
                            PathOp::ClosePath,
                        ],
                        stroke,
                        fill: fill_color,
                    }));
                }
                oxidocs_core::layout::LayoutContent::ClipStart | oxidocs_core::layout::LayoutContent::ClipEnd => {}
                oxidocs_core::layout::LayoutContent::PresetShape { .. } => {}
                // R-05a: balloon variants not yet rendered to PDF; skip.
                oxidocs_core::layout::LayoutContent::Balloon { .. } => {}
                oxidocs_core::layout::LayoutContent::BalloonConnector { .. } => {}
                // S1120 custGeom paths: not drawn in the wasm PDF writer yet
                oxidocs_core::layout::LayoutContent::VectorPath { .. } => {}
            }
        }

        Page {
            width,
            height,
            media_box: Rectangle { llx: 0.0, lly: 0.0, urx: width, ury: height },
            crop_box: None,
            contents,
            rotation: 0,
        }
    }).collect::<Vec<Page>>();

    // Check if any text span uses CJK fonts
    let has_gothic = pages.iter().any(|p| p.contents.iter().any(|c| {
        matches!(c, ContentElement::Text(t) if t.font_name == "OxiCJK-Gothic")
    }));
    let has_mincho = pages.iter().any(|p| p.contents.iter().any(|c| {
        matches!(c, ContentElement::Text(t) if t.font_name == "OxiCJK-Mincho")
    }));

    let mut embedded_fonts = std::collections::HashMap::new();
    if has_gothic || has_mincho {
        let (gothic, mincho) = get_cjk_fonts();
        if has_gothic {
            embedded_fonts.insert("OxiCJK-Gothic".to_string(), gothic.clone());
        }
        if has_mincho {
            embedded_fonts.insert("OxiCJK-Mincho".to_string(), mincho.clone());
        }
    }
    for (key, face) in get_cjk_proportional_fonts() {
        if pages.iter().any(|p| p.contents.iter().any(|c| {
            matches!(c, ContentElement::Text(t) if t.font_name == *key)
        })) {
            embedded_fonts.insert(key.to_string(), face.clone());
        }
    }

    // Embed only the bundled Latin faces this document actually names, so a
    // one-page memo does not carry five families' worth of outlines.
    let mut used_latin: std::collections::BTreeSet<&str> = std::collections::BTreeSet::new();
    for page in &pages {
        for element in &page.contents {
            if let ContentElement::Text(span) = element {
                if let Some((name, _)) = oxidocs_fonts::latin_faces()
                    .into_iter()
                    .find(|(name, _)| *name == span.font_name.as_str())
                {
                    used_latin.insert(name);
                }
            }
        }
    }
    for name in used_latin {
        if let Some(face) = latin_face(name) {
            embedded_fonts.insert(name.to_string(), face.clone());
        }
    }

    for element in layout.pages.iter().flat_map(|page| &page.elements) {
        if let oxidocs_core::layout::LayoutContent::Text { font_family, bold, italic, text, .. } = &element.content {
            if text.is_empty() { continue; }
            let family = font_family.as_deref().unwrap_or("Helvetica");
            let key = crate::fonts::resource_name(family, *bold, *italic);
            if !embedded_fonts.contains_key(&key) {
                if let Some(font) = crate::fonts::embedded_font(family, *bold, *italic) {
                    embedded_fonts.insert(key, font);
                }
            }
        }
    }

    // Cut each embedded face down to the characters drawn with it, as the CLI
    // export does. Whole faces made a one-page Japanese quotation 7.8 MB (all
    // of OxiMincho) against 28 KB from the CLI. A face that is also drawn by
    // glyph index keeps every glyph, since those indices are the original ones.
    let mut drawn: std::collections::HashMap<&str, std::collections::BTreeSet<char>> = std::collections::HashMap::new();
    let mut by_glyph: std::collections::HashSet<&str> = std::collections::HashSet::new();
    for element in pages.iter().flat_map(|p| &p.contents) {
        match element {
            ContentElement::Text(span) => {
                drawn.entry(span.font_name.as_str()).or_default().extend(span.text.chars());
            }
            ContentElement::GlyphRun(run) => {
                by_glyph.insert(run.span.font_name.as_str());
            }
            _ => {}
        }
    }
    for (key, font) in embedded_fonts.iter_mut() {
        if by_glyph.contains(key.as_str()) {
            continue;
        }
        if let Some(chars) = drawn.get(key.as_str()) {
            if let Some(subset) = subset_embedded_font(font, chars) {
                *font = subset;
            }
        }
    }

    PdfDocument {
        version: PdfVersion::new(1, 7),
        info: DocumentInfo {
            title,
            producer: Some("Oxi (docx→PDF)".to_string()),
            ..Default::default()
        },
        pages,
        outline: Vec::new(),
        embedded_fonts,
    }
}

/// A copy of a TrueType face holding only the glyphs of `chars`, renumbered,
/// with the cmap and widths to match. None when the face is not TrueType,
/// draws none of the characters, or the subsetter fails; the caller then keeps
/// the whole face.
pub fn subset_embedded_font(
    font: &oxipdf_core::ir::EmbeddedFont,
    chars: &std::collections::BTreeSet<char>,
) -> Option<oxipdf_core::ir::EmbeddedFont> {
    use oxipdf_core::font_util;
    use oxipdf_core::ir::{EmbeddedFont, FontFormat};
    if !matches!(font.format, FontFormat::TrueType) || font.data.starts_with(b"OTTO") {
        return None;
    }
    let cmap = font_util::parse_cmap_table(&font.data);
    let widths = font_util::parse_ttf_widths(&font.data);
    let mut remapper = subsetter::GlyphRemapper::new();
    let mut unicode_to_gid = std::collections::HashMap::new();
    let mut cid_widths = std::collections::HashMap::new();
    // A character the face lacks is drawn as glyph 0; keep that glyph's own
    // width, or the viewer falls back to a full em and pushes the rest of the
    // span along.
    if let Some(&w) = widths.get(&0) {
        cid_widths.insert(0, w);
    }
    for &c in chars {
        let Some(&gid) = cmap.get(&(c as u32)) else { continue };
        if gid == 0 {
            continue;
        }
        let new = remapper.remap(gid);
        unicode_to_gid.insert(c as u32, new);
        if let Some(&w) = widths.get(&gid) {
            cid_widths.insert(new, w);
        }
    }
    if unicode_to_gid.is_empty() {
        return None;
    }
    let data = subsetter::subset(&font.data, 0, &remapper).ok()?;
    if data.starts_with(b"OTTO") {
        return None;
    }
    Some(EmbeddedFont {
        ps_name: font.ps_name.clone().or_else(|| font_util::parse_ps_name(&font.data)),
        data,
        format: FontFormat::TrueType,
        unicode_to_gid,
        cid_widths,
    })
}

pub fn resolve_wasm_font(base_font: &str, bold: bool, italic: bool, text: &str) -> String {
    // Check if font is a known CJK family
    let is_cjk_font = matches!(base_font,
        "ＭＳ ゴシック" | "MS ゴシック" | "MS Gothic" | "ＭＳ Ｐゴシック" | "MS PGothic"
        | "Yu Gothic" | "游ゴシック" | "メイリオ" | "Meiryo" | "ヒラギノ角ゴ"
        | "ＭＳ 明朝" | "MS 明朝" | "MS Mincho" | "ＭＳ Ｐ明朝" | "MS PMincho"
        | "Yu Mincho" | "游明朝" | "ヒラギノ明朝"
    );

    if is_cjk_font {
        return match oxidocs_fonts::cjk_stand_in(base_font).0 {
            "OxiMincho" => "OxiCJK-Mincho".to_string(),
            _ => "OxiCJK-Gothic".to_string(),
        };
    }

    // Check if text contains CJK characters (fallback)
    let has_cjk = text.chars().any(|c| c as u32 > 0x2E7F);
    if has_cjk {
        return "OxiCJK-Gothic".to_string();
    }

    // A bundled stand-in carries real outlines and a real width table, so it
    // beats naming a base-14 font and letting the viewer decide.
    if let Some(family) = latin_stand_in(base_font) {
        return match (bold, italic) {
            (false, false) => family.to_string(),
            (true, false) => format!("{family}-B"),
            (false, true) => format!("{family}-I"),
            (true, true) => format!("{family}-BI"),
        };
    }

    // Unknown family: keep the old base-14 naming, which at least names a font
    // every viewer has.
    if bold {
        match base_font {
            "Helvetica" => "Helvetica-Bold".to_string(),
            _ => base_font.to_string(),
        }
    } else {
        base_font.to_string()
    }
}

/// Parse "#RRGGBB" hex color to PDF Color.
pub fn parse_hex_color(hex: &str) -> Option<oxipdf_core::ir::Color> {
    let hex = hex.strip_prefix('#').unwrap_or(hex);
    if hex.len() != 6 {
        return None;
    }
    let r = u8::from_str_radix(&hex[0..2], 16).ok()? as f64 / 255.0;
    let g = u8::from_str_radix(&hex[2..4], 16).ok()? as f64 / 255.0;
    let b = u8::from_str_radix(&hex[4..6], 16).ok()? as f64 / 255.0;
    Some(oxipdf_core::ir::Color::Rgb(r, g, b))
}

/// Parse and lay out a .docx and draw its pages: the document the PDF export
/// writes and the native viewers draw tiles of. Fails when a glyph a page
/// selects is not in its registered font program.
pub fn docx_pages(data: &[u8], title: Option<String>) -> Result<oxipdf_core::ir::PdfDocument, String> {
    let doc = oxidocs_core::parse_docx(data).map_err(|e| e.to_string())?;
    let engine = oxidocs_core::layout::LayoutEngine::for_document(&doc);
    let layout = engine.layout(&doc);
    crate::fonts::validate_layout_glyphs(&layout)?;
    Ok(layout_to_pdf(&layout, &doc, title))
}
