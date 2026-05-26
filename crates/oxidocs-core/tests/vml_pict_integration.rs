//! Integration tests: parse `<w:pict>` (VML legacy shape) end-to-end
//! and verify `Page.shapes` Shape.* fields after `parse_docx`.
//!
//! `parse_vml_pict` at parser/ooxml.rs:3712 handles the LEGACY VML
//! shape format (<v:rect>, <v:shape>, <v:oval>, <v:roundrect>,
//! <v:line>) that pre-DrawingML docs use. Distinct entry point from
//! parse_drawing (S316 for anchored, S292 for inline).
//!
//! Parser code path tested:
//!   - [parser/ooxml.rs:3712](crates/oxidocs-core/src/parser/ooxml.rs#L3712)
//!     `parse_vml_pict` end-to-end on five distinct shape types
//!     plus the CSS-length unit converter.
//!
//! Non-obvious behaviors pinned:
//!   - VML shape type mapping (parser/ooxml.rs:3758-3773):
//!     <v:rect> → "rect", <v:oval> → "oval",
//!     <v:roundrect> → "roundRect" (camelCase, NOT "roundrect"),
//!     <v:line> → "line", <v:shape type="#_x0000_t185"/> →
//!     "bracketPair" (special-case for the double-bracket 〔〕 shape
//!     CLAUDE.md S70 2026-04-13). Generic <v:shape> without t185
//!     → "rect" as a fallback.
//!   - CSS-like style attribute parsing — five unit converters in
//!     parse_css_length:
//!     "pt"  → val           (raw)
//!     "in"  → val * 72.0    (1 inch = 72pt)
//!     "cm"  → val * 28.3465
//!     "mm"  → val * 2.83465
//!     "px"  → val * 0.75    (96dpi → 72pt: 72/96)
//!     no suffix → val raw
//!   - Color leading `#` strip: `fillcolor="#FF0000"` →
//!     fill="FF0000" via `.trim_start_matches('#')`. A regression
//!     that lost the strip would surface "#FF0000" downstream and
//!     break hex-comparing consumers.
//!   - Boolean polarity-flip: `filled="f"` OR `filled="false"` →
//!     no_fill=true → Shape.fill stays None (even if fillcolor was
//!     specified). Same idiom for `stroked`. The val=false case
//!     also disables the stroke_width 0.75 default (parser/ooxml.rs:
//!     3895 `if no_stroke { None } else { ... or(Some(0.75)) }`).
//!   - VML absolute-position routing: `style="position:absolute;
//!     margin-left:X;margin-top:Y"` → Shape.position = Some with
//!     h_relative=v_relative="text" HARDCODED (line 3880-3881).
//!     DIFFERENT from DrawingML where these come from positionH/V
//!     `relativeFrom` attrs. A unification refactor must preserve
//!     the divergence.
//!   - stroke_width default = Some(0.75) when stroke is enabled but
//!     no explicit weight (parser/ooxml.rs:3895). NOT None. A
//!     refactor that "cleaned up" this default would silently drop
//!     stroke from every VML shape without explicit weight.
//!   - Shape elements are matched on `Event::Start` ONLY
//!     (parser/ooxml.rs:3758) — self-closing `<v:rect/>` triggers
//!     `Event::Empty` and falls through to the no-op branch. All
//!     fixtures here use explicit start+end pairs to actually
//!     exercise the shape-recognition path. Latent parser
//!     asymmetry: any refactor that "normalizes" Empty to
//!     Start+End in upstream reader config would suddenly start
//!     recognizing self-closing shapes that previously silently
//!     dropped.
//!
//! Fixtures live in `tools/fixtures/vml_pict_samples/` and are
//! authored by `tools/metrics/build_vml_pict_repro_fixtures.py`.

use std::fs;

use oxidocs_core::ir::{Block, Document, Shape};
use oxidocs_core::parse_docx;

fn fixture_path(name: &str) -> std::path::PathBuf {
    let crate_root = std::env::current_dir().unwrap();
    let workspace_root = crate_root.parent().unwrap().parent().unwrap();
    workspace_root
        .join("tools")
        .join("fixtures")
        .join("vml_pict_samples")
        .join(name)
}

fn load(name: &str) -> Option<Document> {
    let path = fixture_path(name);
    if !path.exists() {
        eprintln!("skipping: {} not found", path.display());
        return None;
    }
    let data = fs::read(&path).expect("read fixture");
    Some(parse_docx(&data).expect("parse fixture"))
}

/// Collect shapes from all paragraphs in all pages. VML pict shapes
/// inside a `<w:r>` end up in Paragraph.shapes (per
/// parser/ooxml.rs:1725 ParagraphResult construction), NOT
/// Page.shapes. Page.shapes is only populated by body-level
/// AlternateContent (line 838-841). This helper hides the routing
/// detail from each test.
fn first_paragraph_shapes(doc: &Document) -> &Vec<Shape> {
    for page in &doc.pages {
        for block in &page.blocks {
            if let Block::Paragraph(p) = block {
                if !p.shapes.is_empty() {
                    return &p.shapes;
                }
            }
        }
    }
    panic!("no paragraph with shapes found")
}

#[test]
fn v1_vml_rect_basic_fill_stroke_with_pound_strip_and_anchor() {
    let Some(doc) = load("v1_vml_rect_basic.docx") else { return };
    let shapes = first_paragraph_shapes(&doc);
    assert_eq!(
        shapes.len(),
        1,
        "VML pict shape routes to Paragraph.shapes (NOT Page.shapes — \
         see parser/ooxml.rs:1725 ParagraphResult construction)"
    );

    let s = &shapes[0];
    assert_eq!(s.shape_type, "rect", "<v:rect> → shape_type=\"rect\"");
    assert!((s.width - 100.0).abs() < 0.001, "width:100pt → 100.0");
    assert!((s.height - 50.0).abs() < 0.001, "height:50pt → 50.0");

    // Leading '#' STRIPPED from fillcolor / strokecolor.
    assert_eq!(
        s.fill.as_deref(),
        Some("FF0000"),
        "fillcolor=\"#FF0000\" → fill=\"FF0000\" (# stripped via trim_start_matches)"
    );
    assert_eq!(
        s.stroke_color.as_deref(),
        Some("0000FF"),
        "strokecolor=\"#0000FF\" → stroke_color=\"0000FF\""
    );

    let sw = s.stroke_width.expect("strokeweight populates");
    assert!(
        (sw - 2.0).abs() < 0.001,
        "strokeweight=\"2pt\" → 2.0pt, got {}",
        sw
    );

    assert_eq!(
        s.v_text_anchor.as_deref(),
        Some("middle"),
        "v-text-anchor:middle → v_text_anchor=\"middle\""
    );
}

#[test]
fn v1_vml_bracket_pair_t185_special_case() {
    // <v:shape type="#_x0000_t185"/> → "bracketPair". This is the
    // ONLY path in parse_vml_pict that recognizes a specific
    // VML preset type ID; everything else falls through to "rect".
    // CLAUDE.md S70 2026-04-13 documents this as the double-bracket
    // 〔〕 shape — VML doesn't expose a dedicated <v:bracketPair>
    // element, only the t185 type ID on generic <v:shape>.
    let Some(doc) = load("v1_vml_bracket_pair_t185.docx") else { return };
    let s = &first_paragraph_shapes(&doc)[0];

    assert_eq!(
        s.shape_type, "bracketPair",
        "<v:shape type=\"#_x0000_t185\"/> → \"bracketPair\" \
         (CLAUDE.md S70: double-bracket 〔〕)"
    );
    assert!((s.width - 50.0).abs() < 0.001);
    assert!((s.height - 30.0).abs() < 0.001);
}

#[test]
fn v1_vml_oval_no_fill_no_stroke_polarity_flips_both() {
    let Some(doc) = load("v1_vml_oval_no_fill_no_stroke.docx") else { return };
    let s = &first_paragraph_shapes(&doc)[0];

    assert_eq!(s.shape_type, "oval", "<v:oval> → \"oval\"");

    // filled="f" → no_fill=true → Shape.fill stays None EVEN
    // THOUGH fillcolor="#AAAAAA" was specified. The polarity-flip
    // SUPPRESSES the color storage.
    assert!(
        s.fill.is_none(),
        "filled=\"f\" SUPPRESSES Shape.fill (despite fillcolor=\"#AAAAAA\")"
    );

    // stroked="f" → no_stroke=true → Shape.stroke_color stays None
    // AND Shape.stroke_width stays None (disables the 0.75 default).
    assert!(
        s.stroke_color.is_none(),
        "stroked=\"f\" SUPPRESSES Shape.stroke_color"
    );
    assert!(
        s.stroke_width.is_none(),
        "stroked=\"f\" SUPPRESSES Shape.stroke_width default (NOT Some(0.75))"
    );
}

#[test]
fn v1_vml_roundrect_absolute_position_hardcodes_text_anchor() {
    let Some(doc) = load("v1_vml_roundrect_absolute_position.docx") else { return };
    let s = &first_paragraph_shapes(&doc)[0];

    assert_eq!(
        s.shape_type, "roundRect",
        "<v:roundrect> → \"roundRect\" (camelCase, NOT \"roundrect\")"
    );
    assert!((s.width - 80.0).abs() < 0.001);
    assert!((s.height - 40.0).abs() < 0.001);

    let pos = s
        .position
        .as_ref()
        .expect("position:absolute populates Shape.position");
    assert!(
        (pos.x - 50.0).abs() < 0.001,
        "margin-left:50pt → pos.x=50.0, got {}",
        pos.x
    );
    assert!(
        (pos.y - 30.0).abs() < 0.001,
        "margin-top:30pt → pos.y=30.0, got {}",
        pos.y
    );

    // HARDCODED h_relative=v_relative="text" in VML path
    // (parser/ooxml.rs:3880-3881). DIFFERENT from DrawingML
    // (S316) where these come from positionH/V relativeFrom attrs.
    assert_eq!(
        pos.h_relative.as_deref(),
        Some("text"),
        "VML absolute-position hardcodes h_relative=\"text\" \
         (DIFFERENT from DrawingML where this comes from relativeFrom)"
    );
    assert_eq!(pos.v_relative.as_deref(), Some("text"));
}

#[test]
fn v1_vml_css_units_inch_converts_to_72_points() {
    let Some(doc) = load("v1_vml_css_units.docx") else { return };
    let s = &first_paragraph_shapes(&doc)[0];

    // width:1in → 1.0 * 72.0 = 72.0pt (parse_css_length "in" branch).
    assert!(
        (s.width - 72.0).abs() < 0.001,
        "width:1in → 72.0pt (val * 72.0), got {}",
        s.width
    );
    // height:36pt → 36.0pt (raw).
    assert!(
        (s.height - 36.0).abs() < 0.001,
        "height:36pt → 36.0pt (raw, no conversion), got {}",
        s.height
    );

    // No fillcolor/strokecolor declared, no filled="f"/stroked="f".
    // → stroke_width defaults to Some(0.75) per parser/ooxml.rs:3895.
    // A regression that "cleaned up" this default would silently
    // produce shapes without stroke on every VML shape lacking
    // explicit weight.
    let sw = s
        .stroke_width
        .expect("stroked default ON populates stroke_width=Some(0.75)");
    assert!(
        (sw - 0.75).abs() < 0.001,
        "no explicit strokeweight + stroked-default-on → 0.75pt fallback, got {}",
        sw
    );
}

#[test]
fn all_five_fixtures_produce_exactly_one_shape() {
    let cases: &[&str] = &[
        "v1_vml_rect_basic.docx",
        "v1_vml_bracket_pair_t185.docx",
        "v1_vml_oval_no_fill_no_stroke.docx",
        "v1_vml_roundrect_absolute_position.docx",
        "v1_vml_css_units.docx",
    ];
    for name in cases {
        let path = fixture_path(name);
        if !path.exists() {
            eprintln!("skipping {}", name);
            continue;
        }
        let data = fs::read(&path).unwrap();
        let doc = parse_docx(&data)
            .unwrap_or_else(|e| panic!("failed to parse {}: {:?}", name, e));
        let shapes = first_paragraph_shapes(&doc);
        assert_eq!(
            shapes.len(),
            1,
            "{} should produce exactly 1 shape in Paragraph.shapes",
            name
        );
    }
}
