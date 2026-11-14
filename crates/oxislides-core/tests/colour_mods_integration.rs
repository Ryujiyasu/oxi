// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! Colour modifiers against PowerPoint's own output, byte for byte.
//!
//! `tools/metrics/gen_pptx_lummod_round.py` authors one rectangle per arm
//! (a base `srgbClr` + lumMod / lumOff / lum / satMod / hueOff ... children)
//! and exports a PDF from PowerPoint; the `.txt` beside each deck lists the
//! fill PowerPoint drew for every arm, in shape order.

use oxislides_core::ir::Presentation;

fn fills(p: &Presentation) -> Vec<String> {
    p.slides
        .iter()
        .flat_map(|s| s.shapes.iter())
        .map(|sh| {
            sh.fill_color
                .clone()
                .unwrap_or_default()
                .to_ascii_uppercase()
        })
        .collect()
}

/// (PowerPoint's fill, base, modifiers) per arm.
fn truth(txt: &str) -> Vec<(String, String, String)> {
    txt.lines()
        .filter(|l| !l.starts_with('#') && !l.trim().is_empty())
        .map(|l| {
            let mut it = l.splitn(3, ' ');
            let ppt = it.next().unwrap().to_string();
            let base = it.next().unwrap().to_string();
            (ppt, base, it.next().unwrap_or("").to_string())
        })
        .collect()
}

fn mismatches(deck: &[u8], txt: &str) -> Vec<String> {
    let p = oxislides_core::parser::parse_pptx(deck).expect("deck parses");
    let got = fills(&p);
    let want = truth(txt);
    assert_eq!(got.len(), want.len(), "one shape per arm");
    got.iter()
        .zip(&want)
        .filter(|(g, (w, _, _))| *g != w)
        .map(|(g, (w, base, mods))| format!("{base} {mods}: oxi {g} ppt {w}"))
        .collect()
}

/// 416 arms: 16 bases x lumMod / lumOff / lum / satMod / hueOff / sat. The
/// exact-.5 ties (white under lumMod 50% = 7F7F7F, 808080 under lumOff -10%
/// = 676767) come out PowerPoint's way only because the colour is held as
/// linear light in 1/100000 units on the way in and out.
#[test]
fn lum_and_sat_modifiers_match_powerpoint_byte_for_byte() {
    check(
        "lummod_round",
        include_bytes!("../../../tests/fixtures/pptx_colour_mods/lummod_round.pptx"),
        include_str!("../../../tests/fixtures/pptx_colour_mods/lummod_round.txt"),
    );
}

/// 130 arms: sat / satOff / satMod / hue / hueOff on greys, pure primaries
/// (hue exactly 0) and tints. Saturation is never clamped (satOff -30% on
/// 808080 is 5AA6FF).
#[test]
fn hue_and_sat_modifiers_match_powerpoint_byte_for_byte() {
    check(
        "satmod_hue",
        include_bytes!("../../../tests/fixtures/pptx_colour_mods/satmod_hue.pptx"),
        include_str!("../../../tests/fixtures/pptx_colour_mods/satmod_hue.txt"),
    );
}

/// 192 arms: a grey saturated then turned through the hue wheel in 15° steps,
/// and an absolute hue set before the saturation (it is lost on a grey).
#[test]
fn saturating_a_grey_matches_powerpoint_byte_for_byte() {
    check(
        "grey_hue",
        include_bytes!("../../../tests/fixtures/pptx_colour_mods/grey_hue.pptx"),
        include_str!("../../../tests/fixtures/pptx_colour_mods/grey_hue.txt"),
    );
}

fn check(name: &str, deck: &[u8], txt: &str) {
    let bad = mismatches(deck, txt);
    assert!(
        bad.is_empty(),
        "{name}: {} arms differ:
{}",
        bad.len(),
        bad.join(
            "
"
        )
    );
}
