// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

use super::*;
use crate::ir::{BarPos, FracBarType, MathAlignment, MathBlock, MathExpr};
use std::io::{Cursor, Write};

// Independently authored package parts. No corpus documents, images, metadata
// or font programmes are needed to run these regression tests in a checkout.
fn package(body: &str, settings: &str) -> Vec<u8> {
    let styles = r#"<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="24"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:before="0" w:after="0"/></w:pPr></w:pPrDefault></w:docDefaults></w:styles>"#;

    package_with_styles(body, settings, styles)
}

fn package_with_styles(body: &str, settings: &str, styles: &str) -> Vec<u8> {
    let mut zip = zip::ZipWriter::new(Cursor::new(Vec::new()));
    let options = zip::write::SimpleFileOptions::default();
    let document = format!(r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"><w:body>{body}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:bottom="1440" w:left="1440" w:right="1440"/></w:sectPr></w:body></w:document>"#);
    let settings = format!(r#"<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">{settings}</w:settings>"#);
    let content_types = r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>"#;
    let root_rels = r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="document" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"#;
    let document_rels = r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="styles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="settings" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/></Relationships>"#;
    for (name, text) in [
        ("[Content_Types].xml", content_types),
        ("_rels/.rels", root_rels),
        ("word/document.xml", document.as_str()),
        ("word/_rels/document.xml.rels", document_rels),
        ("word/styles.xml", styles),
        ("word/settings.xml", settings.as_str()),
    ] {
        zip.start_file(name, options).unwrap();
        zip.write_all(text.as_bytes()).unwrap();
    }
    zip.finish().unwrap().into_inner()
}

fn document() -> Document {
    let bytes = package("<w:p><w:r><w:t>SEED</w:t></w:r></w:p>", "");
    let mut doc = crate::parser::parse_docx(&bytes).unwrap();
    doc.pages[0].doc_grid_no_type = true;
    doc
}

fn paragraph(text: &str, size: f32) -> Paragraph {
    let doc = document();
    let Block::Paragraph(mut p) = doc.pages[0].blocks[0].clone() else { panic!("paragraph") };
    p.style = ParagraphStyle::default();
    p.style.space_before = Some(0.0);
    p.style.space_after = Some(0.0);
    p.runs[0].text = text.into();
    p.runs[0].style = RunStyle { font_family: Some("Arial".into()), font_size: Some(size), ..RunStyle::default() };
    p
}

fn table(paragraphs: &[Paragraph], width: f32) -> Table {
    let body = format!("<w:tbl><w:tblPr><w:tblW w:w=\"{}\" w:type=\"dxa\"/></w:tblPr><w:tblGrid><w:gridCol w:w=\"{}\"/></w:tblGrid><w:tr><w:tc><w:p><w:r><w:t>CELL</w:t></w:r></w:p></w:tc></w:tr></w:tbl>", width * 20.0, width * 20.0);
    let doc = crate::parser::parse_docx(&package(&body, "")).unwrap();
    let Block::Table(mut table) = doc.pages[0].blocks[0].clone() else { panic!("table") };
    table.rows[0].cells[0].blocks = paragraphs.iter().cloned().map(Block::Paragraph).collect();
    table
}

fn text_elements(layout: &LayoutResult) -> Vec<&LayoutElement> {
    layout.pages.iter().flat_map(|p| p.elements.iter()).filter(|e|
        matches!(&e.content, LayoutContent::Text { text, .. } if !text.is_empty())).collect()
}

fn text_at<'a>(layout: &'a LayoutResult, label: &str) -> &'a LayoutElement {
    let found: Vec<_> = text_elements(layout).into_iter().filter(|e|
        matches!(&e.content, LayoutContent::Text { text, .. } if text == label)).collect();
    assert_eq!(found.len(), 1, "exactly one source marker {label}");
    found[0]
}

fn source_text(layout: &LayoutResult, paragraph_index: usize) -> String {
    text_elements(layout).into_iter().filter(|e| e.paragraph_index == Some(paragraph_index))
        .filter_map(|e| match &e.content { LayoutContent::Text { text, .. } => Some(text.as_str()), _ => None }).collect()
}

fn normalized(text: &str) -> String { text.chars().filter(|c| !c.is_whitespace()).collect() }

#[test]
fn wrapped_source_survives_exit_from_a_float_band() {
    for height in [20.0, 40.0, 80.0] {
        let mut doc = document();
        let text = "alpha beta gamma delta epsilon ".repeat(24);
        doc.pages[0].blocks = vec![Block::Paragraph(paragraph(&text, 12.0))];
        let image: Image = serde_json::from_value(serde_json::json!({
            "data": [], "width": 180.0, "height": height,
            "position": {"x": 0.0, "y": 0.0, "h_relative": "margin", "v_relative": "paragraph"},
            "wrap_type": "Square"
        })).unwrap();
        doc.pages[0].floating_images.push(image);
        let layout = LayoutEngine::for_document(&doc).layout(&doc);
        assert_eq!(normalized(&source_text(&layout, 0)), normalized(&text));
        for e in text_elements(&layout) {
            assert!(e.x.is_finite() && e.y.is_finite() && e.width > 0.0 && e.height > 0.0);
            assert!(e.x >= 72.0 - 0.01 && e.x + e.width <= 540.0 + 0.01);
        }
    }
}

#[test]
fn full_width_float_preserves_the_first_line_and_following_text() {
    for height in [20.0, 40.0] {
        let mut doc = document();
        doc.pages[0].blocks = vec![Block::Paragraph(paragraph("BEFORE", 12.0)), Block::Paragraph(paragraph("AFTER", 12.0))];
        let image: Image = serde_json::from_value(serde_json::json!({
            "data": [], "width": 468.0, "height": height, "anchor_block_index": 0,
            "position": {"x": 0.0, "y": 0.0, "h_relative": "margin", "v_relative": "paragraph"},
            "wrap_type": "TopAndBottom"
        })).unwrap();
        doc.pages[0].floating_images.push(image);
        let layout = LayoutEngine::for_document(&doc).layout(&doc);
        assert!(text_at(&layout, "BEFORE").y >= 72.0 + height - 0.01);
        assert!(text_at(&layout, "AFTER").y > text_at(&layout, "BEFORE").y);
    }
}

#[test]
fn unequal_columns_retain_all_source_after_a_column_control() {
    for width in [96.0, 144.0, 216.0] {
        let mut doc = document();
        let source = "FIRST\x0BSECOND alpha beta gamma delta epsilon ".to_owned() + &"zeta eta theta ".repeat(8);
        let p = paragraph(&source, 12.0);
        doc.pages[0].blocks = vec![Block::Paragraph(p)];
        doc.pages[0].columns = Some(crate::ir::ColumnLayout {
            num: 2, space: Some(18.0), equal_width: false, separator: false,
            columns: vec![crate::ir::ColumnDef { width, space: Some(18.0) }, crate::ir::ColumnDef { width: 450.0 - width, space: None }],
        });
        let layout = LayoutEngine::for_document(&doc).layout(&doc);
        let rendered = source_text(&layout, 0);
        assert_eq!(normalized(&rendered), normalized(&source.replace('\x0B', "")));
        let second = text_elements(&layout).into_iter().find(|e|
            matches!(&e.content, LayoutContent::Text { text, .. } if text.starts_with("SECOND"))).unwrap();
        assert!((second.x - (72.0 + width + 18.0)).abs() < 0.01);
    }
}

#[test]
fn absolute_float_keeps_each_preceding_box_and_each_row_once() {
    for height in [64.0, 65.0, 66.0, 67.0] {
        for row_count in [1usize, 18] {
            for prefix_table in [false, true] {
                let mut doc = document();
                let mut prefix = paragraph("PREFIX", 12.0);
                prefix.style.line_spacing_rule = Some("exact".into());
                prefix.style.line_spacing = Some(height);
                let prefix = if prefix_table { Block::Table(table(&[prefix], 468.0)) } else { Block::Paragraph(prefix) };
                let mut floating = table(&[paragraph("ROW00", 12.0)], 468.0);
                let original_row = floating.rows[0].clone();
                floating.rows = (0..row_count).map(|i| {
                    let mut row = original_row.clone();
                    let Block::Paragraph(p) = &mut row.cells[0].blocks[0] else { panic!("row") };
                    p.runs[0].text = format!("ROW{i:02}");
                    row
                }).collect();
                floating.style.position = Some(serde_json::from_value(serde_json::json!({
                    "x": 72.0, "y": 138.0, "h_anchor": "page", "v_anchor": "page"
                })).unwrap());
                doc.pages[0].blocks = vec![prefix, Block::Table(floating), Block::Paragraph(paragraph("END", 12.0))];
                let layout = LayoutEngine::for_document(&doc).layout(&doc);
                text_at(&layout, "PREFIX"); text_at(&layout, "END");
                for row in 0..row_count { text_at(&layout, &format!("ROW{row:02}")); }
                for e in text_elements(&layout) { assert!(e.y.is_finite() && e.height > 0.0); }
            }
        }
    }
}

#[test]
fn zero_float_offsets_retain_the_same_source_origin_for_each_reference() {
    let mut baseline = None;
    for reference in [None, Some("page"), Some("margin"), Some("text")] {
        let mut doc = document();
        let mut prefix = paragraph("PREFIX", 12.0);
        prefix.style.line_spacing_rule = Some("exact".into()); prefix.style.line_spacing = Some(30.0);
        let mut floating = table(&[paragraph("FLOAT", 12.0)], 468.0);
        floating.style.position = Some(serde_json::from_value(serde_json::json!({
            "x": 0.0, "y": 0.0, "h_anchor": "margin", "v_anchor": reference
        })).unwrap());
        doc.pages[0].blocks = vec![Block::Paragraph(prefix), Block::Table(floating), Block::Paragraph(paragraph("END", 12.0))];
        let layout = LayoutEngine::for_document(&doc).layout(&doc);
        let positions: Vec<_> = ["PREFIX", "FLOAT", "END"].into_iter().map(|label| text_at(&layout, label).y).collect();
        if let Some(previous) = &baseline { assert_eq!(&positions, previous); } else { baseline = Some(positions); }
    }
}

#[test]
fn multibyte_cell_text_does_not_borrow_the_previous_runs_font_size() {
    let mut heights = Vec::new();
    for tail in [".", "\u{2026}"] {
        let mut p = paragraph("LARGE\n", 26.0);
        let mut small = p.runs[0].clone(); small.text = tail.into(); small.style.font_size = Some(10.0);
        p.runs.push(small);
        let doc = document(); let engine = LayoutEngine::for_document(&doc);
        heights.push(engine.estimate_para_height(&p, 261.75, None, None, true, None, None));
    }
    assert!((heights[0] - heights[1]).abs() < 0.01);
    assert!(heights[0] < 60.0, "the second line must use its own 10pt run");
}

#[test]
fn cell_font_baselines_follow_source_runs_when_columns_are_swapped() {
    let mut expected = None;
    for swapped in [false, true] {
        let mut doc = document();
        let mut body = paragraph("BODY", 10.0); body.runs[0].style.font_family = Some("Arial".into());
        let mut symbol = paragraph("SYMBOL", 18.0); symbol.runs[0].style.font_family = Some("Times New Roman".into());
        let mut t = table(&[body], 220.0);
        let mut cell = t.rows[0].cells[0].clone(); cell.blocks = vec![Block::Paragraph(symbol)];
        t.rows[0].cells.push(cell); t.grid_columns = vec![220.0, 220.0];
        if swapped { t.rows[0].cells.swap(0, 1); }
        doc.pages[0].blocks = vec![Block::Table(t)];
        let layout = LayoutEngine::for_document(&doc).layout(&doc);
        let y = [text_at(&layout, "BODY").y, text_at(&layout, "SYMBOL").y];
        if let Some(previous) = expected { assert_eq!(y, previous); } else { expected = Some(y); }
    }
}

#[test]
fn wrapped_cell_lines_do_not_inherit_a_later_large_run() {
    let mut doc = document();
    let small = paragraph("SMALL", 10.0);
    let mut large = small.runs[0].clone(); large.text = "\nLARGE".into(); large.style.font_size = Some(26.0);
    let mut mixed = small.clone(); mixed.runs.push(large);
    doc.pages[0].blocks = vec![Block::Table(table(&[mixed], 261.75))];
    let layout = LayoutEngine::for_document(&doc).layout(&doc);
    let small_text = text_at(&layout, "SMALL"); let large_text = text_at(&layout, "LARGE");
    assert!(large_text.y > small_text.y);
    assert!(large_text.height > small_text.height);
    assert!(matches!(&small_text.content, LayoutContent::Text { font_size, .. } if *font_size == 10.0));
    assert!(matches!(&large_text.content, LayoutContent::Text { font_size, .. } if *font_size == 26.0));
}

#[test]
fn hanging_cell_tabs_are_positioned_controls_with_a_retained_following_span() {
    // Request a first usable stop at 36pt. A hanging position before it is
    // itself a stop, so cover coincidence and an earlier explicit stop.
    for indent in [36.0, 72.0] {
        let mut doc = document();
        let mut p = paragraph("\tAFTER", 10.0);
        p.style.indent_left = Some(indent);
        p.style.indent_first_line = Some(-indent);
        p.style.tab_stops = vec![crate::ir::TabStop { position: 36.0, alignment: crate::ir::TabStopAlignment::Left, leader: None, clear: false }];
        doc.pages[0].blocks = vec![Block::Table(table(&[p], 261.75))];
        let layout = LayoutEngine::for_document(&doc).layout(&doc);
        let tab = text_at(&layout, "\t"); let after = text_at(&layout, "AFTER");
        assert!(tab.width > 0.0 && (tab.width - 36.0).abs() < 0.01,
            "indent={indent}, tab width={}", tab.width);
        assert!((after.x - (tab.x + tab.width)).abs() < 0.01);
        assert!((after.y - tab.y).abs() < 0.01);
    }
}

fn nested_fraction() -> MathExpr {
    MathExpr::Fraction {
        num: Box::new(MathExpr::Text("1".into())),
        den: Box::new(MathExpr::Fraction {
            num: Box::new(MathExpr::Text("1".into())),
            den: Box::new(MathExpr::Subscript { base: Box::new(MathExpr::Text("n".into())), sub: Box::new(MathExpr::Text("1".into())) }),
            bar_type: FracBarType::Bar,
        }), bar_type: FracBarType::Bar,
    }
}

fn display(expr: MathExpr, reduce: bool) -> MathBlock {
    MathBlock::Display { content: vec![expr], reduce_fraction_size: reduce, jc: MathAlignment::Left, host: None }
}

fn baseline(e: &LayoutElement) -> f32 { e.y + e.height * (2.0 / 3.0) }

#[test]
fn nested_fraction_preserves_full_size_and_compact_baseline_gaps() {
    for (size, gap) in [(12.0, 13.08), (18.0, 19.56)] {
        let (elements, _) = math::emit_math_block(&display(nested_fraction(), false), 0.0, 0.0, size);
        let numerator = elements.iter().filter(|e| matches!(&e.content, LayoutContent::Text { text, .. } if text == "1")).nth(1).unwrap();
        let denominator = elements.iter().find(|e| matches!(&e.content, LayoutContent::Text { text, .. } if text == "\u{1d45b}")).unwrap();
        assert!((baseline(denominator) - baseline(numerator) - gap).abs() <= 0.15);
        assert!(matches!(&numerator.content, LayoutContent::Text { font_size, .. } if (*font_size - size).abs() < 0.01));
    }
}

#[test]
fn nested_fraction_script_sizes_retain_nominal_half_points() {
    for (base, script, second) in [(8.0,5.5,4.5),(10.0,7.0,6.0),(10.5,7.5,6.0),(11.0,8.0,6.5),(13.0,9.0,7.5),(14.0,10.0,8.0),(16.0,11.5,9.5),(20.0,14.5,12.0)] {
        let (elements, _) = math::emit_math_block(&display(nested_fraction(), true), 0.0, 0.0, base);
        let sizes: Vec<_> = elements.iter().filter_map(|e| match &e.content { LayoutContent::Text { text, font_size, .. } if text == "1" => Some(*font_size), _ => None }).collect();
        assert_eq!(sizes.len(), 3);
        for (actual, expected) in sizes.into_iter().zip([base, script, second]) { assert!((actual - expected).abs() < 0.01); }
    }
}

#[test]
fn fraction_numerator_overbar_preserves_the_baseline_gap() {
    let expr = MathExpr::Seq(vec![MathExpr::Text("=".into()), MathExpr::Fraction {
        num: Box::new(MathExpr::Bar { pos: BarPos::Top, base: Box::new(MathExpr::Text("X".into())) }),
        den: Box::new(MathExpr::Text("n".into())), bar_type: FracBarType::Bar,
    }]);
    let (elements, _) = math::emit_math_block(&display(expr, false), 0.0, 0.0, 12.0);
    let outer = elements.iter().find(|e| matches!(&e.content, LayoutContent::Text { text, .. } if text == "=")).unwrap();
    let numerator = elements.iter().find(|e| matches!(&e.content, LayoutContent::Text { text, .. } if text == "\u{1d44b}")).unwrap();
    assert!((baseline(outer) - baseline(numerator) - 9.12).abs() <= 0.12);
}

#[test]
fn display_math_spacing_and_package_settings_do_not_leak_between_parses() {
    let equation = "<w:p><m:oMathPara><m:oMath><m:f><m:num><m:r><m:t>1</m:t></m:r></m:num><m:den><m:r><m:t>n</m:t></m:r></m:den></m:f></m:oMath></m:oMathPara></w:p>";
    for reduce in [true, false, true, false] {
        let settings = format!("<m:mathPr><m:smallFrac m:val=\"{}\"/></m:mathPr>", if reduce { "on" } else { "off" });
        let parsed = crate::parser::parse_docx(&package(equation, &settings)).unwrap();
        assert!(matches!(&parsed.pages[0].blocks[0], Block::Math(MathBlock::Display { reduce_fraction_size, .. }) if *reduce_fraction_size == reduce));
    }
    let mut origins = Vec::new();
    for after in [0.0, 10.0] {
        let mut doc = document();
        let mut block = display(nested_fraction(), false);
        let MathBlock::Display { host, .. } = &mut block else { unreachable!() };
        let mut style = paragraph("", 12.0).style; style.space_after = Some(after); *host = Some(Box::new(style));
        doc.pages[0].blocks = vec![Block::Paragraph(paragraph("M1", 12.0)), Block::Math(block), Block::Paragraph(paragraph("M2", 12.0))];
        let layout = LayoutEngine::for_document(&doc).layout(&doc);
        origins.push([text_at(&layout, "M1").y, text_at(&layout, "M2").y]);
    }
    assert_eq!(origins[0][0], origins[1][0]);
    assert!(origins[1][1] > origins[0][1]);
}


#[test]
fn application_font_fallback_keeps_undeclared_defaults_out_of_source_ir() {
    let body = "<w:p><w:r><w:t>iii WWW 0123</w:t></w:r></w:p>";
    let root = r#"<w:style w:type="paragraph" w:styleId="Base" w:default="1"><w:rPr><w:sz w:val="22"/></w:rPr></w:style>"#;
    let styles = format!(r#"<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">{root}</w:styles>"#);
    let doc = crate::parser::parse_docx(&package_with_styles(body, "", &styles)).unwrap();
    assert!(doc.styles.doc_default_run_style.is_none());
    assert!(doc.styles.doc_default_para_style.is_none());
    let source = serde_json::to_value(&doc).unwrap();
    let actual = LayoutEngine::for_document(&doc).layout(&doc);
    assert!(doc.styles.doc_default_run_style.is_none());
    assert_eq!(serde_json::to_value(&doc).unwrap(), source);
    let mut explicit = doc.clone();
    explicit.styles.doc_default_run_style = Some(RunStyle {
        font_family: Some("Times New Roman".into()), ..RunStyle::default()
    });
    let expected = LayoutEngine::for_document(&explicit).layout(&explicit);
    // LayoutResult is a runtime type without Serialize. Compare page geometry,
    // every text fragment's paint properties, and its source identity directly.
    fn snapshot(layout: &LayoutResult) -> serde_json::Value {
        serde_json::Value::Array(layout.pages.iter().map(|page| {
            let elements: Vec<_> = page.elements.iter().map(|element| {
                let LayoutContent::Text {
                    text, font_size, font_family, bold, italic, underline,
                    underline_style, strikethrough, double_strikethrough, color,
                    highlight, field_type, character_spacing, text_scale,
                    is_vertical, effects,
                } = &element.content else { panic!("text-only authored fixture"); };
                serde_json::json!({
                    "position": [element.x, element.y, element.width, element.height],
                    "text": text, "size": font_size, "face": font_family,
                    "bold": bold, "italic": italic, "underline": underline,
                    "underline_style": underline_style, "strike": strikethrough,
                    "double_strike": double_strikethrough, "color": color,
                    "highlight": highlight, "field": field_type,
                    "spacing": character_spacing, "scale": text_scale,
                    "vertical": is_vertical, "effects": [effects.shadow, effects.emboss, effects.imprint, effects.outline, effects.no_fill],
                    "paragraph": element.paragraph_index, "run": element.run_index,
                    "offset": element.char_offset, "source_text": element.source_text,
                    "source_len": element.source_char_len,
                    "container": element.source_container_index,
                    "extent": element.source_paragraph_extent,
                    "prefix": element.source_paragraph_prefix,
                    "baseline": element.baseline_offset, "text_y_off": element.text_y_off,
                    "clip": element.horizontal_clip,
                })
            }).collect();
            serde_json::json!({"width": page.width, "height": page.height, "elements": elements})
        }).collect())
    }
    assert_eq!(snapshot(&actual), snapshot(&expected));

    for (language, declared_face, expected_face) in [
        ("ja-JP", None, None),
        ("en-US", None, Some("Times New Roman")),
        ("ja-JP", Some("Arial"), Some("Arial")),
    ] {
        let face = declared_face.map(|name| format!(r#"<w:rFonts w:ascii="{name}"/>"#)).unwrap_or_default();
        let styles = format!(r#"<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:docDefaults><w:rPrDefault><w:rPr><w:lang w:val="{language}"/>{face}</w:rPr></w:rPrDefault></w:docDefaults>{root}</w:styles>"#);
        let doc = crate::parser::parse_docx(&package_with_styles(body, "", &styles)).unwrap();
        assert_eq!(doc.styles.doc_default_run_style.as_ref().unwrap().font_family.as_deref(), declared_face);
        assert_eq!(LayoutEngine::for_document(&doc).default_font_family.as_deref(), expected_face);
    }
    let styles = r#"<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:style w:type="paragraph" w:styleId="Other"><w:rPr><w:sz w:val="22"/></w:rPr></w:style></w:styles>"#;
    let doc = crate::parser::parse_docx(&package_with_styles(body, "", styles)).unwrap();
    assert!(doc.styles.doc_default_run_style.is_none());
    assert!(LayoutEngine::for_document(&doc).default_font_family.is_none());
}


#[test]
fn row_vertical_margin_overrides_preserve_geometry_and_page_breaks() {
    fn authored(rows: usize, row_override: bool, mixed_cell_margin: bool) -> Document {
        let mut doc = document();
        let mut grid = table(&[paragraph("SEED", 7.5)], 180.0);
        grid.grid_columns = vec![90.0, 90.0];
        grid.style.default_cell_margins = Some(CellMargins {
            top: Some(if row_override { 3.5 } else { 0.0 }),
            bottom: Some(if row_override { 2.85 } else { 0.0 }),
            left: Some(0.0), right: Some(0.0),
        });
        let seed = grid.rows[0].clone();
        grid.rows = (0..rows).map(|i| {
            let mut row = seed.clone();
            row.cell_margins_override = row_override.then_some(CellMargins {
                top: Some(0.0), bottom: Some(0.0), left: None, right: None,
            });
            row.cells = (0..2).map(|col| {
                let mut cell = seed.cells[0].clone();
                cell.blocks = vec![Block::Paragraph(paragraph(&format!("ROW{i}COL{col}"), 7.5))];
                cell.margins = (mixed_cell_margin && col == 0).then_some(CellMargins {
                    top: Some(2.0), bottom: None, left: None, right: None,
                });
                cell
            }).collect();
            row
        }).collect();
        doc.pages[0].blocks = vec![Block::Table(grid), Block::Paragraph(paragraph("AFTER", 7.5))];
        doc
    }
    fn geometry(layout: &LayoutResult) -> Vec<(usize, String, f32, f32, f32, f32)> {
        layout.pages.iter().enumerate().flat_map(|(page, p)| p.elements.iter().filter_map(move |e| {
            match &e.content {
                LayoutContent::Text { text, .. } if !text.is_empty() =>
                    Some((page, text.clone(), e.x, e.y, e.width, e.height)),
                _ => None,
            }
        })).collect()
    }
    for rows in [2, 80] {
        for mixed_cell_margin in [false, true] {
            let expected_doc = authored(rows, false, mixed_cell_margin);
            let actual_doc = authored(rows, true, mixed_cell_margin);
            let expected = LayoutEngine::for_document(&expected_doc).layout(&expected_doc);
            let actual = LayoutEngine::for_document(&actual_doc).layout(&actual_doc);
            assert_eq!(actual.pages.len(), expected.pages.len());
            if rows == 80 { assert!(actual.pages.len() > 1); }
            assert_eq!(geometry(&actual), geometry(&expected),
                "rows={rows}, mixed_cell_margin={mixed_cell_margin}");
        }
    }
}
