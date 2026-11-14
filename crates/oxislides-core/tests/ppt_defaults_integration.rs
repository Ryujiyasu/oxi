// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! A deck authored by PowerPoint itself through its UI defaults
//! (`tools/metrics/gen_pptx_ppt_defaults.py`): inserted shapes that colour
//! themselves through `p:style`, theme colours with a brightness (`lumMod` /
//! `lumOff`), connectors, a new table on the default table style, the title
//! layout's subtitle, and an overfull body PowerPoint shrank (`normAutofit`).
//! Every expected colour below is read from PowerPoint's own PDF export of
//! the same file.

use oxislides_core::ir::{ShapeContent, SlideBullet};

fn deck() -> oxislides_core::ir::Presentation {
    let data = include_bytes!("../../../tests/fixtures/ppt_defaults.pptx");
    oxislides_core::parser::parse_pptx(data).expect("deck parses")
}

fn paras(sh: &oxislides_core::ir::Shape) -> &Vec<oxislides_core::ir::SlideParagraph> {
    match &sh.content {
        ShapeContent::TextBox { paragraphs } | ShapeContent::AutoShape { paragraphs } => paragraphs,
        _ => panic!("not a text shape"),
    }
}

fn first_text(sh: &oxislides_core::ir::Shape) -> String {
    match &sh.content {
        ShapeContent::TextBox { paragraphs } | ShapeContent::AutoShape { paragraphs } => paragraphs
            .first()
            .and_then(|p| p.runs.first())
            .map(|r| r.text.clone())
            .unwrap_or_default(),
        _ => String::new(),
    }
}

/// Byte for byte, ties included (PowerPoint's 50% grey is 7F7F7F; see
/// colour_mods_integration.rs).
fn near(a: &str, b: &str) -> bool {
    a.eq_ignore_ascii_case(b)
}

#[test]
fn inserted_shapes_take_fill_line_and_text_colour_from_their_style() {
    let p = deck();
    let s = &p.slides[0];
    let rect = s.shapes.iter().find(|sh| first_text(sh) == "Default rectangle").unwrap();
    assert_eq!(rect.fill_color.as_deref(), Some("156082"));
    assert_eq!(rect.border_color.as_deref(), Some("042433"));
    assert_eq!(rect.border_width, Some(1.5));
    assert_eq!(paras(rect)[0].runs[0].color.as_deref(), Some("FFFFFF"));
}

#[test]
fn theme_colours_with_a_brightness_are_resolved() {
    let p = deck();
    let s = &p.slides[0];
    let colour = |t: &str| {
        let sh = s.shapes.iter().find(|sh| first_text(sh) == t).unwrap();
        paras(sh)[0].runs[0].color.clone().unwrap()
    };
    assert!(near(&colour("Text 1, lighter 50%"), "7F7F7F"));
    assert!(near(&colour("Background 1, darker 50%"), "7F7F7F"));
    assert_eq!(colour("Accent 1, darker 50%"), "0B3041");
    assert_eq!(colour("Accent 2, lighter 40%"), "F2AA84");
    let light = s.shapes.iter().find(|sh| first_text(sh) == "Light fill").unwrap();
    assert_eq!(light.fill_color.as_deref(), Some("C1E5F5"));
}

#[test]
fn connectors_take_their_line_from_the_style() {
    let p = deck();
    let conns: Vec<_> = p.slides[0]
        .shapes
        .iter()
        .filter(|sh| sh.shape_type.as_deref().is_some_and(|t| t.contains("Connector")))
        .collect();
    assert_eq!(conns.len(), 2);
    for c in conns {
        assert_eq!(c.border_color.as_deref(), Some("156082"));
        assert!(c.border_width.is_some());
    }
}

#[test]
fn default_table_style_paints_header_bands_and_rules() {
    let p = deck();
    let table = p.slides[1]
        .shapes
        .iter()
        .find_map(|sh| match &sh.content {
            ShapeContent::Table { table } => Some(table),
            _ => None,
        })
        .unwrap();
    let head = &table.rows[0][0];
    assert_eq!(head.fill_color.as_deref(), Some("156082"));
    let r = &head.paragraphs[0].runs[0];
    assert_eq!((r.color.as_deref(), r.bold), (Some("FFFFFF"), Some(true)));
    let bottom = head.borders[3].as_ref().unwrap();
    assert_eq!((bottom.color.as_str(), bottom.width), ("FFFFFF", 3.0));
    assert_eq!(table.rows[1][0].fill_color.as_deref(), Some("CCD2D8"));
    assert_eq!(table.rows[2][0].fill_color.as_deref(), Some("E7EAED"));
    assert_eq!(table.rows[1][0].paragraphs[0].runs[0].color.as_deref(), Some("000000"));
}

#[test]
fn subtitle_takes_no_bullet_and_no_hanging_indent() {
    let p = deck();
    let sub = p.slides[2].shapes.iter().find(|sh| sh.ph_type.as_deref() == Some("subTitle")).unwrap();
    let m = oxislides_core::layout::resolve_level(&p.master_styles.body, &sub.ph_levels, 0);
    assert!(matches!(m.bullet, SlideBullet::None), "{:?}", m.bullet);
    assert_eq!((m.mar_l, m.indent), (0.0, 0.0));
}

#[test]
fn master_body_level_carries_its_pitch_and_colour() {
    let b = &deck().master_styles.body[0];
    assert_eq!(b.line_spacing, Some(0.9));
    assert_eq!(b.spc_bef_pts, Some(10.0));
    assert_eq!(b.color.as_deref(), Some("000000"));
}

#[test]
fn shrunk_body_is_drawn_at_the_rounded_scaled_size() {
    let p = deck();
    let body = p.slides[3].shapes.iter().find(|sh| sh.ph_type.as_deref() == Some("body")).unwrap();
    assert_eq!(body.norm_autofit, Some((0.55, 0.2)));
    for para in paras(body) {
        // 28pt x 55% = 15.4, which PowerPoint draws at 15pt.
        assert_eq!(para.runs[0].font_size, Some(15.0));
        // lnSpc 90% - 20%
        assert!((para.line_spacing.unwrap() - 0.7).abs() < 1e-6);
    }
}
