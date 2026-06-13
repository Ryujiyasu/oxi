// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! Public-API integration tests for oxihanko.
//!
//! Pins the externally-visible behavior of the hanko stamp generator
//! and PDF signing wrapper. The oxihanko crate has 9 in-src unit tests
//! but ZERO integration tests prior to this file — last engine-style
//! crate missing public-API integration coverage.
//!
//! Tests cover: SVG generation determinism, stamp style variants,
//! color and config defaults, and the preview wrapper's pure-function
//! contract (same input → byte-identical output).

use oxihanko::{
    generate_stamp_svg, preview_stamp, StampColor, StampConfig, StampStyle,
};

fn base_config(name: &str) -> StampConfig {
    StampConfig {
        name: name.into(),
        style: StampStyle::Round,
        color: StampColor::vermilion(),
        size: 100,
        border_width: 0.08,
        font_ratio: 0.45,
        date: None,
    }
}

#[test]
fn generate_round_stamp_produces_svg() {
    let svg = generate_stamp_svg(&base_config("山田"));
    assert!(svg.starts_with("<svg"), "expected <svg prefix, got {:?}", &svg[..40.min(svg.len())]);
    assert!(svg.contains("</svg>"), "expected </svg> closing tag");
    assert!(svg.contains("山田"), "stamp name must appear in SVG output");
}

#[test]
fn generate_square_stamp_produces_svg() {
    let mut cfg = base_config("会社印");
    cfg.style = StampStyle::Square;
    let svg = generate_stamp_svg(&cfg);
    assert!(svg.starts_with("<svg"));
    // Square stamps split chars across grid cells, each in its own
    // <text> element; "会社印" appears as separate 会 / 社 / 印 chars.
    assert!(svg.contains("会"), "stamp name char must appear in SVG");
    assert!(svg.contains("社"));
    assert!(svg.contains("印"));
}

#[test]
fn generate_oval_stamp_produces_svg() {
    let mut cfg = base_config("銀行");
    cfg.style = StampStyle::Oval;
    let svg = generate_stamp_svg(&cfg);
    assert!(svg.starts_with("<svg"));
    assert!(svg.contains("銀行"));
}

#[test]
fn stamp_generation_is_deterministic() {
    // Same config → byte-identical SVG. Required for reproducible PDF
    // signatures and content-addressed stamp caches.
    let cfg = base_config("山田");
    let a = generate_stamp_svg(&cfg);
    let b = generate_stamp_svg(&cfg);
    assert_eq!(a, b, "generate_stamp_svg must be deterministic for identical config");
}

#[test]
fn preview_stamp_matches_generate() {
    // preview_stamp is documented as "no signing"; it must produce the
    // same SVG as generate_stamp_svg for the same config.
    let cfg = base_config("山田");
    assert_eq!(generate_stamp_svg(&cfg), preview_stamp(&cfg));
}

#[test]
fn different_styles_produce_different_svgs() {
    // Round vs Square vs Oval must produce visually distinguishable output.
    let mut round_cfg = base_config("印");
    round_cfg.style = StampStyle::Round;
    let mut square_cfg = base_config("印");
    square_cfg.style = StampStyle::Square;
    let mut oval_cfg = base_config("印");
    oval_cfg.style = StampStyle::Oval;
    let r = generate_stamp_svg(&round_cfg);
    let s = generate_stamp_svg(&square_cfg);
    let o = generate_stamp_svg(&oval_cfg);
    assert_ne!(r, s, "Round and Square must produce different SVGs");
    assert_ne!(s, o, "Square and Oval must produce different SVGs");
    assert_ne!(r, o, "Round and Oval must produce different SVGs");
}

#[test]
fn color_palette_constructors_produce_distinct_colors() {
    let v = StampColor::vermilion();
    let r = StampColor::red();
    let k = StampColor::black();
    assert_ne!((v.r, v.g, v.b), (r.r, r.g, r.b), "vermilion ≠ red");
    assert_ne!((v.r, v.g, v.b), (k.r, k.g, k.b), "vermilion ≠ black");
    assert_ne!((r.r, r.g, r.b), (k.r, k.g, k.b), "red ≠ black");
    assert_eq!(k.r + k.g + k.b, 0, "black must be #000000");
}

#[test]
fn stamp_config_default_uses_round_vermilion() {
    let d = StampConfig::default();
    assert!(matches!(d.style, StampStyle::Round));
    let v = StampColor::vermilion();
    assert_eq!((d.color.r, d.color.g, d.color.b), (v.r, v.g, v.b));
    assert_eq!(d.size, 100);
    assert!(d.date.is_none());
}

#[test]
fn stamp_color_changes_appear_in_svg() {
    let mut cfg = base_config("印");
    cfg.color = StampColor::red();
    let red_svg = generate_stamp_svg(&cfg);
    cfg.color = StampColor::black();
    let black_svg = generate_stamp_svg(&cfg);
    assert_ne!(red_svg, black_svg, "color change must propagate to SVG");
}

#[test]
fn date_field_appears_in_svg_when_set() {
    let mut cfg = base_config("印");
    let svg_no_date = generate_stamp_svg(&cfg);
    cfg.date = Some("2026.05.28".to_string());
    let svg_with_date = generate_stamp_svg(&cfg);
    assert_ne!(svg_no_date, svg_with_date, "date toggle must change output");
    assert!(svg_with_date.contains("2026.05.28"), "date string must appear in SVG");
}

#[test]
fn empty_name_does_not_panic() {
    // Edge case: empty name. SVG should still generate without panic.
    let cfg = StampConfig::default();  // name = ""
    let svg = generate_stamp_svg(&cfg);
    assert!(svg.starts_with("<svg"));
}
