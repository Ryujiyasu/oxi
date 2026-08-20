// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! Integration tests for oxislides-core shape & run property parsing.
//!
//! S351's public_api_integration covered parse error paths, slide
//! dimensions, Slide.index, and PptxEditor round-trip. This suite pins the
//! VISUAL property extraction that the renderer depends on: shape geometry
//! (EMU→pt), run bold/italic/color/font, and paragraph alignment. A
//! regression here silently corrupts slide rendering.

use oxislides_core::ir::{GeomCmd, ShapeContent, SlideAlignment};
use oxislides_core::parser::parse_pptx;

const PPTX: &[u8] = include_bytes!("../../../tests/fixtures/basic_test.pptx");

fn pres() -> oxislides_core::ir::Presentation {
    parse_pptx(PPTX).expect("basic_test.pptx must parse")
}

#[test]
fn slide_has_title_and_body_shapes() {
    let p = pres();
    let slide = &p.slides[0];
    assert!(slide.shapes.len() >= 2, "title + body");
}

#[test]
fn shape_geometry_emu_to_points() {
    // basic_test slide1: title shape off=(457200,274638) ext=(8229600,1143000).
    // 1 point = 12700 EMU. → x=36pt, y≈21.63pt, w=648pt, h=90pt.
    let p = pres();
    let title = &p.slides[0].shapes[0];
    assert!((title.x - 36.0).abs() < 0.5, "title x≈36pt, got {}", title.x);
    assert!((title.y - 21.63).abs() < 0.5, "title y≈21.6pt, got {}", title.y);
    assert!((title.width - 648.0).abs() < 1.0, "title w≈648pt, got {}", title.width);
    assert!((title.height - 90.0).abs() < 1.0, "title h≈90pt, got {}", title.height);
}

#[test]
fn shape_geometry_is_finite_and_nonneg() {
    // Every shape must have finite, non-negative geometry (renderer contract).
    let p = pres();
    for (si, slide) in p.slides.iter().enumerate() {
        for (shi, sh) in slide.shapes.iter().enumerate() {
            for (name, v) in [("x", sh.x), ("y", sh.y), ("w", sh.width), ("h", sh.height)] {
                assert!(v.is_finite(), "slide {si} shape {shi} {name} not finite");
                assert!(v >= 0.0, "slide {si} shape {shi} {name} negative: {v}");
            }
        }
    }
}

#[test]
fn title_run_is_bold() {
    let p = pres();
    if let ShapeContent::TextBox { paragraphs } = &p.slides[0].shapes[0].content {
        let run = &paragraphs[0].runs[0];
        assert_eq!(run.text, "Welcome to Oxi");
        assert!(run.bold, "title run must be bold");
    } else {
        panic!("title must be a TextBox");
    }
}

#[test]
fn body_has_italic_colored_run() {
    // basic_test body (shape 1) has an italic run colored 4472C4.
    let p = pres();
    if let ShapeContent::TextBox { paragraphs } = &p.slides[0].shapes[1].content {
        let mut found = false;
        for para in paragraphs {
            for run in &para.runs {
                if run.italic && run.color.as_deref() == Some("4472C4") {
                    found = true;
                }
            }
        }
        assert!(found, "body must contain an italic run colored 4472C4");
    } else {
        panic!("body must be a TextBox");
    }
}

#[test]
fn run_color_is_six_hex_digits_when_present() {
    // Any parsed run color must be a 6-hex-digit string (no leading '#',
    // uppercase or lowercase) — the renderer assumes this format.
    let p = pres();
    for slide in &p.slides {
        for shape in &slide.shapes {
            if let ShapeContent::TextBox { paragraphs } = &shape.content {
                for para in paragraphs {
                    for run in &para.runs {
                        if let Some(c) = &run.color {
                            assert_eq!(c.len(), 6, "color must be 6 hex digits, got {c:?}");
                            assert!(
                                c.chars().all(|ch| ch.is_ascii_hexdigit()),
                                "color must be hex, got {c:?}"
                            );
                        }
                    }
                }
            }
        }
    }
}

#[test]
fn paragraph_alignment_parsed() {
    // basic_test slide1 has a center-aligned paragraph (algn="ctr").
    let p = pres();
    let mut saw_center = false;
    for shape in &p.slides[0].shapes {
        if let ShapeContent::TextBox { paragraphs } = &shape.content {
            for para in paragraphs {
                if matches!(para.alignment, Some(SlideAlignment::Center)) {
                    saw_center = true;
                }
            }
        }
    }
    assert!(saw_center, "slide1 must have a center-aligned paragraph");
}

#[test]
fn alignment_is_none_or_a_valid_variant() {
    // multi_slide.pptx slide 1 title is left/default-aligned. Since S6 an
    // unspecified algn parses to None (inherit the master txStyles level)
    // rather than a materialized Left, so the pin is "None or a real variant",
    // never a panic / garbage.
    let multi = parse_pptx(include_bytes!("../../../tests/fixtures/multi_slide.pptx"))
        .expect("multi_slide.pptx must parse");
    // The exhaustive match guards against an unparsed state and fails to
    // compile if a variant is added without revisiting this test.
    for slide in &multi.slides {
        for shape in &slide.shapes {
            if let ShapeContent::TextBox { paragraphs } = &shape.content {
                for para in paragraphs {
                    match para.alignment {
                        None
                        | Some(
                            SlideAlignment::Left
                            | SlideAlignment::Center
                            | SlideAlignment::Right
                            | SlideAlignment::Justify,
                        ) => {}
                    }
                }
            }
        }
    }
}

// ---------------------------------------------------------------------------
// a:custGeom — explicit outline paths (S-CUSTGEOM, 2026-08-17)
// ---------------------------------------------------------------------------

const CUSTGEOM: &[u8] = include_bytes!("../../../tests/fixtures/custgeom_test.pptx");

fn custgeom_shape() -> oxislides_core::ir::Shape {
    let p = parse_pptx(CUSTGEOM).expect("custgeom_test.pptx must parse");
    p.slides[0]
        .shapes
        .iter()
        .find(|s| s.custom_geometry.is_some())
        .expect("the fixture carries one custGeom shape")
        .clone()
}

#[test]
fn custgeom_path_space_and_commands_survive_parsing() {
    let shape = custgeom_shape();
    let geom = shape.custom_geometry.unwrap();
    assert!(!geom.unsupported, "the fixture uses only modelled commands");
    assert_eq!(geom.paths.len(), 1, "one a:path, as in every corpus custGeom");
    let path = &geom.paths[0];
    assert_eq!((path.w, path.h), (100.0, 200.0), "a:path @w/@h is the local space");
    assert!(!path.fill_none);
    // moveTo, lnTo, cubicBezTo, lnTo, close — the whole modelled vocabulary.
    assert_eq!(path.commands.len(), 5);
    assert!(matches!(path.commands[0], GeomCmd::MoveTo(50.0, 0.0)));
    assert!(matches!(path.commands[1], GeomCmd::LineTo(100.0, 100.0)));
    assert!(matches!(
        path.commands[2],
        GeomCmd::CubicTo(90.0, 150.0, 60.0, 200.0, 50.0, 200.0)
    ));
    assert!(matches!(path.commands[4], GeomCmd::Close));
}

#[test]
fn custgeom_shape_keeps_its_box_and_fill() {
    // The outline replaces the shape's rectangle only for DRAWING; the box and
    // the fill colour a consumer paints it with are unchanged.
    let shape = custgeom_shape();
    assert!((shape.x - 100.0).abs() < 0.01, "1270000 EMU = 100pt, got {}", shape.x);
    assert!((shape.width - 100.0).abs() < 0.01);
    assert!((shape.height - 50.0).abs() < 0.01);
    assert_eq!(shape.fill_color.as_deref(), Some("C00000"));
    assert!(shape.shape_type.is_none(), "custGeom shapes carry no prstGeom");
}

#[test]
fn preset_shapes_carry_no_custom_geometry() {
    // basic_test's shapes are placeholders/presets: none must acquire a
    // geometry, or the rectangular fallback would be skipped for them.
    let p = pres();
    for slide in &p.slides {
        for shape in &slide.shapes {
            assert!(shape.custom_geometry.is_none());
        }
    }
}

#[test]
fn pictures_rotate_with_their_shape_by_default() {
    // `a:blipFill/@rotWithShape` is absent on every fixture shape, and every
    // one of the dev corpus's 2141 shape-level blipFills declares "1", so the
    // default the renderer relies on is "rotate". A silent flip of this default
    // would leave 489 rotated corpus images upright.
    for pptx in [CUSTGEOM, PPTX] {
        let p = parse_pptx(pptx).expect("fixture must parse");
        for slide in &p.slides {
            for shape in &slide.shapes {
                assert!(shape.rot_with_shape, "default must be rotate-with-shape");
            }
        }
    }
}

// --- a:endParaRPr ---------------------------------------------------------
// An empty paragraph takes its line height from its paragraph mark, so the
// mark's size has to survive parsing. `emptypara_test.pptx` carries both XML
// shapes of the element: `<a:endParaRPr sz="2400"/>` self-closing, and one
// with an `a:solidFill` child. quick-xml routes those to different events, and
// a handler present in only one arm silently drops half the corpus.

const EMPTYPARA_PPTX: &[u8] = include_bytes!("../../../tests/fixtures/emptypara_test.pptx");

fn emptypara_paragraphs() -> Vec<oxislides_core::ir::SlideParagraph> {
    let p = parse_pptx(EMPTYPARA_PPTX).expect("emptypara_test.pptx must parse");
    p.slides[0]
        .shapes
        .iter()
        .find_map(|s| match &s.content {
            ShapeContent::AutoShape { paragraphs } if paragraphs.len() >= 4 => {
                Some(paragraphs.clone())
            }
            _ => None,
        })
        .expect("the text box with four paragraphs")
}

#[test]
fn end_para_size_read_from_self_closing_element() {
    let paras = emptypara_paragraphs();
    assert_eq!(
        paras[1].end_para_size,
        Some(24.0),
        "<a:endParaRPr sz=\"2400\"/> arrives as Event::Empty"
    );
}

#[test]
fn end_para_size_read_when_the_element_has_children() {
    let paras = emptypara_paragraphs();
    assert_eq!(
        paras[2].end_para_size,
        Some(10.0),
        "an endParaRPr wrapping a:solidFill arrives as Event::Start"
    );
}

#[test]
fn end_para_size_does_not_leak_between_paragraphs() {
    let paras = emptypara_paragraphs();
    assert_eq!(paras[0].end_para_size, None, "AAA declares no paragraph mark size");
    assert_eq!(paras[3].end_para_size, None, "and the mark size is reset per paragraph");
}

// --- a:gradFill stop alpha ------------------------------------------------
// `<a:alpha val="20000"/>` has nothing but its attribute, so it ALWAYS arrives
// as Event::Empty. A stop-alpha handler written only into the Start arm reads
// nothing and every translucent ramp paints opaque: d15's illustrations are
// white at 20% fading to 0 over a purple slide, and painting them solid put a
// white slab across the artwork.

const GRADALPHA_PPTX: &[u8] = include_bytes!("../../../tests/fixtures/gradalpha_test.pptx");

#[test]
fn gradient_stop_alpha_is_read_from_the_self_closing_element() {
    let p = parse_pptx(GRADALPHA_PPTX).expect("gradalpha_test.pptx must parse");
    let g = p.slides[0]
        .shapes
        .iter()
        .find_map(|s| s.gradient.as_ref())
        .expect("the shape carries an a:gradFill");
    assert_eq!(g.stops.len(), 3, "three stops");
    assert!(
        (g.stops[0].alpha - 0.2).abs() < 1e-6,
        "first stop is 20% opaque, got {}",
        g.stops[0].alpha
    );
    assert!(
        g.stops[1].alpha.abs() < 1e-6 && g.stops[2].alpha.abs() < 1e-6,
        "the ramp fades to fully transparent, got {} / {}",
        g.stops[1].alpha,
        g.stops[2].alpha
    );
}

// --- a:rPr/a:highlight ----------------------------------------------------
// The highlight carries a colour element of exactly the shape `a:solidFill`
// carries, so a parser that does not track the container reads it as the run's
// TEXT colour. d11 slide 38's "and many more..." is white on dk1 and Oxi drew
// it dk1 with no box. The colour element is Event::Empty when self-closing and
// Event::Start when it wraps a modifier, and both arms have to route it.

const HIGHLIGHT_PPTX: &[u8] = include_bytes!("../../../tests/fixtures/highlight_test.pptx");

fn highlight_runs() -> Vec<oxislides_core::ir::SlideRun> {
    let p = parse_pptx(HIGHLIGHT_PPTX).expect("highlight_test.pptx must parse");
    p.slides[0]
        .shapes
        .iter()
        .find_map(|s| match &s.content {
            ShapeContent::AutoShape { paragraphs } if !paragraphs.is_empty() => {
                Some(paragraphs[0].runs.clone())
            }
            _ => None,
        })
        .expect("the text box's one paragraph")
}

#[test]
fn highlight_is_read_from_both_quick_xml_arms() {
    let runs = highlight_runs();
    assert_eq!(runs.len(), 3, "three runs");
    assert_eq!(runs[0].highlight, None, "the first run asks for no highlight");
    assert_eq!(
        runs[1].highlight.as_deref(),
        Some("FF0000"),
        "a self-closing a:srgbClr arrives as Event::Empty"
    );
    assert_eq!(
        runs[2].highlight.as_deref(),
        Some("00FF00"),
        "an a:srgbClr wrapping a:lumMod arrives as Event::Start"
    );
}

#[test]
fn a_highlight_does_not_become_the_run_colour() {
    let runs = highlight_runs();
    assert_eq!(runs[0].color.as_deref(), Some("112233"));
    assert_eq!(
        runs[1].color.as_deref(),
        Some("FFFFFF"),
        "white text on a red highlight stays white"
    );
    assert_eq!(runs[2].color.as_deref(), Some("000000"));
}

// --- a placeholder whose idx matches nothing -------------------------------
// d24 slide 22's body is `<p:ph idx="4294967295" type="body"/>` -- the sentinel
// PowerPoint writes for an unset idx. Its layout has no body placeholder and
// the master's is idx="1", so an exact-key lookup finds nothing and the size
// falls all the way to the master's `p:txStyles/p:bodyStyle`. PowerPoint drew
// that paragraph at exactly 24.00pt, the master PLACEHOLDER's sz, not the 14pt
// txStyles says.

const PHANYIDX_PPTX: &[u8] = include_bytes!("../../../tests/fixtures/phanyidx_test.pptx");

#[test]
fn a_placeholder_with_an_unmatched_idx_still_inherits_the_master_level() {
    let p = parse_pptx(PHANYIDX_PPTX).expect("phanyidx_test.pptx must parse");
    let shape = p.slides[0]
        .shapes
        .iter()
        .find(|s| matches!(&s.content, ShapeContent::AutoShape { paragraphs }
                           if !paragraphs.is_empty()))
        .expect("the body placeholder");
    let lvl = shape.ph_levels.first().expect("the master body level");
    assert_eq!(
        lvl.font_size,
        Some(24.0),
        "the master's body PLACEHOLDER declares 24pt; its txStyles say 14pt"
    );
    assert_eq!(lvl.font_family.as_deref(), Some("Arial"));
}

#[test]
fn title_and_ctr_title_name_the_same_slot() {
    // The slide asks for ctrTitle, the master declares title. 74 title /
    // ctrTitle placeholders over 8 dev decks are reachable only through that
    // alias, d35's "BIG CONCEPT" among them.
    let p = parse_pptx(PHANYIDX_PPTX).expect("phanyidx_test.pptx must parse");
    let shape = p.slides[0]
        .shapes
        .iter()
        .find(|s| matches!(&s.content, ShapeContent::AutoShape { paragraphs }
                           if !paragraphs.is_empty()))
        .expect("the ctrTitle placeholder");
    assert!(
        !shape.ph_levels.is_empty(),
        "a ctrTitle must find the master's title level"
    );
}

#[test]
fn a_level_can_declare_the_highlight() {
    let p = parse_pptx(PHANYIDX_PPTX).expect("phanyidx_test.pptx must parse");
    let shape = p.slides[0]
        .shapes
        .iter()
        .find(|s| matches!(&s.content, ShapeContent::AutoShape { paragraphs }
                           if !paragraphs.is_empty()))
        .expect("the ctrTitle placeholder");
    let lvl = shape.ph_levels.first().expect("the master title level");
    assert_eq!(
        lvl.highlight.as_deref(),
        Some("FFFF00"),
        "a:highlight inside a level's defRPr is the level's box, not its text colour"
    );
    assert_eq!(
        lvl.color.as_deref(),
        Some("112233"),
        "and the level's own solidFill still reads as the text colour"
    );
}

#[test]
fn a_level_can_ask_for_italic() {
    // d16's layout body level is `<a:defRPr i="1" sz="3600"/>` and PowerPoint
    // sets the whole quotation slanted; 18 levels over two dev decks declare
    // one, and MasterStyleLevel had nowhere to put it.
    let p = parse_pptx(PHANYIDX_PPTX).expect("phanyidx_test.pptx must parse");
    let shape = p.slides[0]
        .shapes
        .iter()
        .find(|s| matches!(&s.content, ShapeContent::AutoShape { paragraphs }
                           if !paragraphs.is_empty()))
        .expect("the ctrTitle placeholder");
    assert!(
        shape.ph_levels.first().expect("the master title level").italic,
        "a level's defRPr @i is the level's slant"
    );
}

// --- a rotation inherited through a mirror ---------------------------------
// Two nested groups, each rot=-90 flipH=1. Accumulating the two component-wise
// gives -180 with the flips cancelling; the real composition is the IDENTITY,
// because a mirror reverses the rotation it passes through:
// R(-90) F R(-90) F = R(-90) R(+90) F F. d19's layout stacks 29 pencils exactly
// this way and every one of them came out upside down, blunt end up.

const GRPFLIPROT_PPTX: &[u8] = include_bytes!("../../../tests/fixtures/grpfliprot_test.pptx");

#[test]
fn a_group_rotation_reverses_when_it_passes_through_a_mirror() {
    let p = parse_pptx(GRPFLIPROT_PPTX).expect("grpfliprot_test.pptx must parse");
    let leaf = p.slides[0]
        .shapes
        .iter()
        .find(|s| s.fill_color.as_deref() == Some("FF0000"))
        .expect("the layout's leaf shape");
    assert!(
        leaf.rotation.abs() < 0.01,
        "two -90 flipH groups compose to the identity, not {}",
        leaf.rotation
    );
    assert!(!leaf.flip_h, "the two mirrors cancel");
    assert!(!leaf.flip_v);
}
