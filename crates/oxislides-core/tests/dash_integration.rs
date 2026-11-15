// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! A deck PowerPoint authored itself (`tools/metrics/gen_pptx_dash.py`): every
//! `prstDash` preset on lines and outlines, inside a group, on connectors and
//! with round caps. PowerPoint writes `<a:noFill/>` for an unfilled outline and
//! keeps the fill it does NOT show in `spPr/a:extLst` (`a14:hiddenFill`).

const PRESETS: [&str; 10] = [
    "dot", "dash", "lgDash", "dashDot", "lgDashDot", "lgDashDotDot", "sysDash", "sysDot",
    "sysDashDot", "sysDashDotDot",
];

fn deck() -> oxislides_core::ir::Presentation {
    let data = include_bytes!("../../../tests/fixtures/pptx_dash.pptx");
    oxislides_core::parser::parse_pptx(data).expect("deck parses")
}

#[test]
fn every_preset_is_read_in_order() {
    let p = deck();
    for slide in &p.slides[..4] {
        let dashes: Vec<_> = slide.shapes.iter().map(|s| s.border_dash.as_deref()).collect();
        let mut want: Vec<Option<&str>> = PRESETS.iter().map(|s| Some(*s)).collect();
        want.push(None); // `solid`
        assert_eq!(dashes, want);
    }
}

#[test]
fn hidden_fill_in_the_extension_list_is_not_the_fill() {
    let p = deck();
    // slide 4: unfilled rectangles and ellipses; slide 5: two of them in a group
    for slide in &p.slides[3..5] {
        for sh in &slide.shapes {
            assert_eq!(sh.fill_color, None, "{:?}", sh.shape_type);
            if let Some(c) = &sh.border_color {
                assert_eq!(c, "000000");
            }
        }
    }
}

#[test]
fn round_caps_and_group_member_dashes_are_kept() {
    let p = deck();
    let s = &p.slides[4];
    let dashes: Vec<_> = s.shapes.iter().map(|sh| sh.border_dash.as_deref()).collect();
    assert_eq!(
        dashes,
        vec![
            Some("dot"),
            Some("lgDashDot"),
            Some("sysDashDot"),
            Some("dashDot"),
            Some("lgDash"),
            Some("dot"),
            Some("sysDot"),
            Some("lgDashDot")
        ]
    );
    let caps: Vec<_> = s.shapes.iter().map(|sh| sh.line_cap.as_deref()).collect();
    assert_eq!(caps[5], Some("rnd"));
    assert_eq!(caps[6], Some("rnd"));
}
