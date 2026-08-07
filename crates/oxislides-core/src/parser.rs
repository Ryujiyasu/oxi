// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

use std::collections::HashMap;

use quick_xml::events::Event;
use quick_xml::reader::Reader;
use thiserror::Error;

use oxidocs_common::archive::OoxmlArchive;
use oxidocs_common::relationships::parse_relationships;
use oxidocs_common::xml_utils::{emu_to_pt, get_attr, local_name};

use crate::ir::{
    default_chart_bar_dir, default_chart_grouping, default_chart_type, Chart, ChartSeries,
    MasterStyleLevel, MasterTxStyles, Presentation, Shape, ShapeContent, Slide,
    SlideAlignment, SlideBullet, SlideParagraph, SlideRun, Table, TableCell,
};

#[derive(Error, Debug)]
pub enum PptxError {
    #[error("Archive error: {0}")]
    Archive(#[from] oxidocs_common::OxiError),

    #[error("XML error: {0}")]
    Xml(#[from] quick_xml::Error),

    #[error("Invalid data: {0}")]
    InvalidData(String),
}

/// Information about a slide from presentation.xml
struct SlideInfo {
    r_id: String,
}

/// Parse themeN.xml for the minor/major latin typefaces.
///
/// A run with no explicit `a:latin`/`a:ea` font resolves to the theme minor
/// font (body text / textboxes / body placeholders); a TITLE placeholder
/// resolves to the theme MAJOR font. Falls back to "Calibri" (PowerPoint's
/// default theme font) when the theme part is absent or has no latin face.
/// Parse ppt/theme/themeN.xml: the minor/major latin typefaces plus the
/// `<a:clrScheme>` colour map (Spec #10). Returns (minor_font, major_font,
/// scheme_slot -> RGB-hex). A clrScheme entry carries a 6-digit hex in
/// `<a:srgbClr w:val="RRGGBB"/>` (or `<a:sysClr w:lastClr="RRGGBB"/>` for the
/// system slots dk1/lt1) under the slot element (dk1/dk2/lt1/lt2/tx1/tx2/
/// accent1..accent6/hlink/folHlink).
fn parse_theme(xml: &str) -> Result<(String, String, HashMap<String, String>), PptxError> {
    let mut reader = Reader::from_str(xml);
    let mut minor = "Calibri".to_string();
    let mut major = "Calibri".to_string();
    let mut colors: HashMap<String, String> = HashMap::new();
    let mut in_minor = false;
    let mut in_major = false;
    // clrScheme slot tracking: the active slot name (e.g. "accent1") whose
    // child srgbClr/sysClr supplies the hex.
    let mut in_clr_scheme = false;
    let mut cur_slot: Option<String> = None;

    loop {
        match reader.read_event()? {
            Event::Start(e) | Event::Empty(e) => {
                let name = local_name(e.name().as_ref());
                match name.as_str() {
                    "minorFont" => in_minor = true,
                    "majorFont" => in_major = true,
                    "clrScheme" => in_clr_scheme = true,
                    "latin" if in_minor => {
                        if let Some(t) = get_attr(&e, "typeface") {
                            if !t.is_empty() {
                                minor = t;
                            }
                        }
                    }
                    "latin" if in_major => {
                        if let Some(t) = get_attr(&e, "typeface") {
                            if !t.is_empty() {
                                major = t;
                            }
                        }
                    }
                    "dk1" | "dk2" | "lt1" | "lt2" | "tx1" | "tx2" | "accent1" | "accent2"
                    | "accent3" | "accent4" | "accent5" | "accent6" | "hlink" | "folHlink"
                        if in_clr_scheme =>
                    {
                        cur_slot = Some(name);
                    }
                    "srgbClr" if in_clr_scheme && cur_slot.is_some() => {
                        if let Some(v) = get_attr(&e, "val") {
                            if !v.is_empty() {
                                colors.insert(cur_slot.clone().unwrap(), v);
                            }
                        }
                    }
                    "sysClr" if in_clr_scheme && cur_slot.is_some() => {
                        if let Some(v) = get_attr(&e, "lastClr") {
                            if !v.is_empty() {
                                colors.insert(cur_slot.clone().unwrap(), v);
                            }
                        }
                    }
                    _ => {}
                }
            }
            Event::End(e) => {
                let name = local_name(e.name().as_ref());
                match name.as_str() {
                    "minorFont" => in_minor = false,
                    "majorFont" => in_major = false,
                    "clrScheme" => in_clr_scheme = false,
                    "dk1" | "dk2" | "lt1" | "lt2" | "tx1" | "tx2" | "accent1" | "accent2"
                    | "accent3" | "accent4" | "accent5" | "accent6" | "hlink" | "folHlink"
                        if in_clr_scheme =>
                    {
                        cur_slot = None;
                    }
                    _ => {}
                }
            }
            Event::Eof => break,
            _ => {}
        }
    }

    Ok((minor, major, colors))
}

/// True for master outline-level property names: a:lvl1pPr .. a:lvl9pPr.
fn is_master_lvl(name: &str) -> bool {
    name.starts_with("lvl") && name.ends_with("pPr")
}

/// Parse slideMasterN.xml `p:txStyles` into the per-context, per-outline-level
/// inherited marL / indent / bullet / spcBef (Spec #8).
///
/// PowerPoint slides inherit their list formatting from the master's text
/// styles: placeholder BODY text uses `bodyStyle` levels, plain textboxes use
/// `otherStyle`, title placeholders use `titleStyle`. Each Vec is indexed by
/// the 0-based outline level (a:lvlNpPr with N = level+1). A level that is
/// absent from the part stays an empty `MasterStyleLevel` (marL=0, indent=0,
/// bullet=Inherit, spcBef=None), which reproduces a plain paragraph.
fn parse_master_txstyles(xml: &str) -> Result<MasterTxStyles, PptxError> {
    let mut reader = Reader::from_str(xml);
    let mut body: Vec<MasterStyleLevel> = Vec::new();
    let mut other: Vec<MasterStyleLevel> = Vec::new();
    let mut title: Vec<MasterStyleLevel> = Vec::new();
    let mut in_title_style = false;
    let mut in_body_style = false;
    let mut in_other_style = false;
    // Accumulation for the current a:lvlNpPr.
    let mut cur_level = MasterStyleLevel::default();
    let mut in_level = false;
    let mut in_spc_bef = false;
    let mut level_bullet_font: Option<String> = None;

    // Push the just-finished level into the active context Vec.
    macro_rules! push_level {
        () => {
            if in_body_style {
                body.push(cur_level.clone());
            } else if in_other_style {
                other.push(cur_level.clone());
            } else if in_title_style {
                title.push(cur_level.clone());
            }
        };
    }

    loop {
        match reader.read_event()? {
            Event::Start(e) => {
                let name = local_name(e.name().as_ref());
                match name.as_str() {
                    "titleStyle" => in_title_style = true,
                    "bodyStyle" => in_body_style = true,
                    "otherStyle" => in_other_style = true,
                    n if is_master_lvl(n) => {
                        if in_level {
                            // Defensive: a new level opened before the previous
                            // End (malformed) — close the prior one first.
                            push_level!();
                        }
                        cur_level = MasterStyleLevel::default();
                        level_bullet_font = None;
                        if let Some(m) = get_attr(&e, "marL") {
                            if let Ok(v) = m.parse::<f32>() {
                                cur_level.mar_l = emu_to_pt(v);
                            }
                        }
                        if let Some(ix) = get_attr(&e, "indent") {
                            if let Ok(v) = ix.parse::<f32>() {
                                cur_level.indent = emu_to_pt(v);
                            }
                        }
                        // Spec #6: a:lvlNpPr/@algn — inherited horizontal
                        // alignment (titleStyle lvl1pPr algn="ctr" centres
                        // title placeholders).
                        if let Some(algn) = get_attr(&e, "algn") {
                            cur_level.algn = Some(parse_alignment_attr(&algn));
                        }
                        in_level = true;
                        in_spc_bef = false;
                    }
                    "spcBef" if in_level => in_spc_bef = true,
                    // a:defRPr/@sz (hundredths of a point) — the placeholder
                    // default font size for this outline level (Spec #5). The
                    // level's own paragraph run properties override it at
                    // render time; a layout-level txStyles is NOT inherited
                    // (Word render-truth, phfs probe).
                    "defRPr" if in_level => {
                        if let Some(sz) = get_attr(&e, "sz") {
                            if let Ok(v) = sz.parse::<f32>() {
                                cur_level.font_size = Some(v / 100.0);
                            }
                        }
                    }
                    _ => {}
                }
            }
            Event::Empty(e) => {
                let name = local_name(e.name().as_ref());
                match name.as_str() {
                    // Self-closing level: parse and push immediately.
                    n if is_master_lvl(n) => {
                        let mut lvl = MasterStyleLevel::default();
                        if let Some(m) = get_attr(&e, "marL") {
                            if let Ok(v) = m.parse::<f32>() {
                                lvl.mar_l = emu_to_pt(v);
                            }
                        }
                        if let Some(ix) = get_attr(&e, "indent") {
                            if let Ok(v) = ix.parse::<f32>() {
                                lvl.indent = emu_to_pt(v);
                            }
                        }
                        // Spec #6: a:lvlNpPr/@algn on a self-closing level.
                        if let Some(algn) = get_attr(&e, "algn") {
                            lvl.algn = Some(parse_alignment_attr(&algn));
                        }
                        if in_body_style {
                            body.push(lvl);
                        } else if in_other_style {
                            other.push(lvl);
                        } else if in_title_style {
                            title.push(lvl);
                        }
                    }
                    // Children of a non-self-closing level (Start form).
                    "spcPct" if in_spc_bef => {
                        // a:spcBef/a:spcPct/@val is the fraction in 100000ths.
                        if let Some(v) = get_attr(&e, "val") {
                            if let Ok(x) = v.parse::<f32>() {
                                cur_level.spc_bef_pct = Some(x / 100000.0);
                            }
                        }
                    }
                    "buFont" if in_level => {
                        if let Some(tf) = get_attr(&e, "typeface") {
                            if !tf.is_empty() {
                                level_bullet_font = Some(tf);
                            }
                        }
                    }
                    "buChar" if in_level => {
                        if let Some(ch) = get_attr(&e, "char") {
                            if let Some(c) = ch.chars().next() {
                                cur_level.bullet = SlideBullet::Char {
                                    ch: c,
                                    font: level_bullet_font.take(),
                                };
                            }
                        }
                    }
                    "buNone" if in_level => {
                        cur_level.bullet = SlideBullet::None;
                    }
                    "buAutoNum" if in_level => {
                        let kind =
                            get_attr(&e, "type").unwrap_or_else(|| "arabicPeriod".to_string());
                        let start_at = get_attr(&e, "startAt").and_then(|v| v.parse::<u32>().ok());
                        cur_level.bullet = SlideBullet::AutoNum { kind, start_at };
                    }
                    _ => {}
                }
            }
            Event::End(e) => {
                let name = local_name(e.name().as_ref());
                match name.as_str() {
                    "titleStyle" => in_title_style = false,
                    "bodyStyle" => in_body_style = false,
                    "otherStyle" => in_other_style = false,
                    "spcBef" => in_spc_bef = false,
                    n if is_master_lvl(n) && in_level => {
                        push_level!();
                        cur_level = MasterStyleLevel::default();
                        in_level = false;
                        level_bullet_font = None;
                    }
                    _ => {}
                }
            }
            Event::Eof => break,
            _ => {}
        }
    }

    Ok(MasterTxStyles {
        body,
        other,
        title,
    })
}

/// Parse presentation.xml to get slide relationship IDs (in order).
fn parse_presentation_slides(xml: &str) -> Result<(Vec<SlideInfo>, f32, f32), PptxError> {
    let mut reader = Reader::from_str(xml);
    let mut slides = Vec::new();
    // Default slide size: 10 inches x 7.5 inches in EMU
    let mut width_emu: f32 = 9144000.0;
    let mut height_emu: f32 = 6858000.0;

    loop {
        match reader.read_event()? {
            Event::Start(e) | Event::Empty(e) => {
                let name = local_name(e.name().as_ref());
                match name.as_str() {
                    "sldIdLst" => {} // container for sldId entries
                    "sldId" => {
                        // r:id attribute (namespaced, so try raw "r:id" first)
                        let r_id = {
                            let mut found = None;
                            for attr in e.attributes().flatten() {
                                let key =
                                    std::str::from_utf8(attr.key.as_ref()).unwrap_or("");
                                if key == "r:id" {
                                    found = Some(
                                        String::from_utf8_lossy(&attr.value).to_string(),
                                    );
                                    break;
                                }
                            }
                            found.unwrap_or_default()
                        };
                        if !r_id.is_empty() {
                            slides.push(SlideInfo { r_id });
                        }
                    }
                    "sldSz" => {
                        if let Some(cx) = get_attr(&e, "cx") {
                            if let Ok(v) = cx.parse::<f32>() {
                                width_emu = v;
                            }
                        }
                        if let Some(cy) = get_attr(&e, "cy") {
                            if let Ok(v) = cy.parse::<f32>() {
                                height_emu = v;
                            }
                        }
                    }
                    _ => {}
                }
            }
            Event::Eof => break,
            _ => {}
        }
    }

    Ok((slides, emu_to_pt(width_emu), emu_to_pt(height_emu)))
}

/// Spec #6: parse a:a:pPr/@algn (or a:defRPr/paragraph alignment) attribute.
/// "l" (default) / "ctr" / "r" / "just"; any unknown value is treated as Left.
fn parse_alignment_attr(a: &str) -> SlideAlignment {
    match a {
        "ctr" => SlideAlignment::Center,
        "r" => SlideAlignment::Right,
        "just" => SlideAlignment::Justify,
        _ => SlideAlignment::Left,
    }
}

/// Parse a single slide XML into shapes.
fn parse_slide(
    xml: &str,
    slide_index: usize,
    archive: &mut OoxmlArchive,
    slide_rels_path: &str,
    master_ph_geoms: &HashMap<(Option<String>, Option<String>), (f32, f32, f32, f32)>,
    master_ph_anchors: &HashMap<(Option<String>, Option<String>), String>,
    theme_colors: &HashMap<String, String>,
) -> Result<Slide, PptxError> {
    // Parse slide relationships for image resolution
    let rels = if let Ok(Some(rels_xml)) = archive.try_read_part(slide_rels_path) {
        parse_relationships(&rels_xml).unwrap_or_default()
    } else {
        Default::default()
    };

    // Spec #3: build a (ph_type, ph_idx) -> (x,y,w,h) geometry map from the
    // referenced slideLayout's placeholders. A slide placeholder with NO explicit
    // xfrm in its spPr inherits the layout placeholder's geometry. Spec #8
    // follow-up: when the layout placeholder also lacks an explicit xfrm, fall
    // back to the slideMaster's placeholder geometry (the master ph carries the
    // authoritative xfrm for layout-less placeholder slots).
    let (layout_ph_geoms, layout_ph_anchors): (
        HashMap<(Option<String>, Option<String>), (f32, f32, f32, f32)>,
        HashMap<(Option<String>, Option<String>), String>,
    ) = {
        let mut geoms = HashMap::new();
        let mut anchors = HashMap::new();
        for rel in rels.values() {
            if rel.rel_type.ends_with("/slideLayout") {
                let layout_path =
                    resolve_slide_relative_path(slide_rels_path, &rel.target);
                if let Ok(Some(layout_xml)) = archive.try_read_part(&layout_path) {
                    let (g, a) = parse_layout_ph_info(&layout_xml).unwrap_or_default();
                    geoms = g;
                    anchors = a;
                }
                break;
            }
        }
        (geoms, anchors)
    };

    let mut reader = Reader::from_str(xml);
    let mut shapes = Vec::new();
    let mut _depth = 0u32;
    let mut in_sp_tree = false;

    // Slide background state
    let mut in_bg = false;
    let mut in_bg_pr = false;
    let mut slide_background_color: Option<String> = None;

    // Shape state
    let mut in_shape = false;
    let mut shape_x: f32 = 0.0;
    let mut shape_y: f32 = 0.0;
    let mut shape_w: f32 = 0.0;
    let mut shape_h: f32 = 0.0;
    let mut shape_rotation: f32 = 0.0;
    let mut shape_prst: Option<String> = None;
    let mut shape_paragraphs: Vec<SlideParagraph> = Vec::new();
    let mut shape_is_image = false;
    let mut shape_image_r_id: Option<String> = None;
    let mut shape_fill_color: Option<String> = None;
    let mut shape_border_color: Option<String> = None;
    let mut shape_border_width: Option<f32> = None;
    // Placeholder identity (p:ph type/idx from nvPr) and whether spPr had an
    // explicit xfrm. Spec #3: a placeholder without an explicit xfrm inherits
    // its geometry from the referenced slideLayout's matching placeholder.
    let mut shape_ph_type: Option<String> = None;
    let mut shape_ph_idx: Option<String> = None;
    let mut shape_has_xfrm = false;
    // Spec #6: vertical text-anchor from the shape's own a:bodyPr/@anchor
    // (resolved through the placeholder chain at shape end).
    let mut shape_anchor: Option<String> = None;

    // Shape property context tracking
    let mut in_sp_pr = false; // inside <p:spPr> or <xdr:spPr>
    let mut in_ln = false;    // inside <a:ln> (line/border properties)

    // Text-area insets (a:bodyPr lIns/rIns/tIns/bIns, EMU -> pt).
    // Placeholder defaults: 7.2 / 7.2 / 3.6 / 3.6 (Spec #8 measurement).
    let mut in_body_pr = false;
    let mut shape_l_ins: f32 = 7.2;
    let mut shape_r_ins: f32 = 7.2;
    let mut shape_t_ins: f32 = 3.6;
    let mut shape_b_ins: f32 = 3.6;

    // Paragraph state
    let mut in_paragraph = false;
    let mut para_runs: Vec<SlideRun> = Vec::new();
    // Spec #6: paragraph alignment is Option — None = not specified on the
    // paragraph (the master txStyles level alignment applies at render time).
    let mut para_alignment: Option<SlideAlignment> = None;
    // Spec #4: paragraph spacing (a:pPr/a:lnSpc, a:spcBef, a:spcAft)
    let mut para_line_spacing: Option<f32> = None;
    let mut para_space_before: Option<f32> = None;
    let mut para_space_after: Option<f32> = None;
    let mut in_ln_spc = false;
    let mut in_spc_bef = false;
    let mut in_spc_aft = false;
    // Spec #8: bullet / indent (a:pPr/@lvl, @marL, @indent + a:buChar/buFont/buNone/buAutoNum)
    let mut para_lvl: u32 = 0;
    let mut para_mar_l: Option<f32> = None;
    let mut para_indent: Option<f32> = None;
    let mut para_bullet = SlideBullet::Inherit;
    let mut para_bullet_font: Option<String> = None;

    // Run state
    let mut in_run = false;
    let mut run_text = String::new();
    let mut run_bold = false;
    let mut run_italic = false;
    let mut run_font_size: Option<f32> = None;
    let mut run_color: Option<String> = None;
    let mut run_font_family: Option<String> = None;

    let mut in_text = false;

    // Table state (a:graphicFrame -> a:tbl)
    let mut in_graphic_frame = false;
    let mut in_table = false;
    let mut tbl_col_widths: Vec<f32> = Vec::new();
    let mut tbl_row_heights: Vec<f32> = Vec::new();
    let mut tbl_rows: Vec<Vec<TableCell>> = Vec::new();
    let mut in_tbl_row = false;
    let mut tbl_cur_row: Vec<TableCell> = Vec::new();
    let mut in_tbl_cell = false;
    let mut tbl_cur_cell_paragraphs: Vec<SlideParagraph> = Vec::new();

    // Chart state (a:graphicFrame -> a:graphicData uri=.../chart -> c:chart)
    let mut in_chart_graphic = false;
    let mut chart_r_id: Option<String> = None;

    loop {
        match reader.read_event()? {
            Event::Start(e) => {
                let name = local_name(e.name().as_ref());
                _depth += 1;

                match name.as_str() {
                    "bg" => {
                        in_bg = true;
                    }
                    "bgPr" if in_bg => {
                        in_bg_pr = true;
                    }
                    "spTree" => {
                        in_sp_tree = true;
                    }
                    "sp" | "pic" if in_sp_tree => {
                        in_shape = true;
                        shape_x = 0.0;
                        shape_y = 0.0;
                        shape_w = 0.0;
                        shape_h = 0.0;
                        shape_rotation = 0.0;
                        shape_prst = None;
                        shape_paragraphs.clear();
                        shape_is_image = name == "pic";
                        shape_image_r_id = None;
                        shape_fill_color = None;
                        shape_border_color = None;
                        shape_border_width = None;
                        shape_ph_type = None;
                        shape_ph_idx = None;
                        shape_has_xfrm = false;
                        in_body_pr = false;
                        shape_l_ins = 7.2;
                        shape_r_ins = 7.2;
                        shape_t_ins = 3.6;
                        shape_b_ins = 3.6;
                        shape_anchor = None;
                    }
                    "graphicFrame" if in_sp_tree => {
                        // A graphicFrame (table/chart/SmartArt). Reuse the shape
                        // geometry state (a:xfrm off/ext feed shape_x/y/w/h).
                        in_shape = true;
                        in_graphic_frame = true;
                        shape_x = 0.0;
                        shape_y = 0.0;
                        shape_w = 0.0;
                        shape_h = 0.0;
                        shape_rotation = 0.0;
                        shape_prst = None;
                        shape_paragraphs.clear();
                        shape_is_image = false;
                        shape_image_r_id = None;
                        shape_fill_color = None;
                        shape_border_color = None;
                        shape_border_width = None;
                        // Table state reset
                        in_table = false;
                        tbl_col_widths.clear();
                        tbl_row_heights.clear();
                        tbl_rows.clear();
                        in_tbl_row = false;
                        tbl_cur_row.clear();
                        in_tbl_cell = false;
                        tbl_cur_cell_paragraphs.clear();
                        // Chart state reset
                        in_chart_graphic = false;
                        chart_r_id = None;
                        shape_ph_type = None;
                        shape_ph_idx = None;
                        shape_has_xfrm = false;
                        in_body_pr = false;
                        shape_l_ins = 7.2;
                        shape_r_ins = 7.2;
                        shape_t_ins = 3.6;
                        shape_b_ins = 3.6;
                        shape_anchor = None;
                    }
                    "tbl" if in_graphic_frame => {
                        in_table = true;
                        tbl_col_widths.clear();
                        tbl_row_heights.clear();
                        tbl_rows.clear();
                    }
                    "graphicData" if in_graphic_frame => {
                        // a:graphicData/@uri discriminates the content type inside
                        // a graphicFrame: chart, table, SmartArt, etc.
                        if let Some(uri) = get_attr(&e, "uri") {
                            if uri == "http://schemas.openxmlformats.org/drawingml/2006/chart" {
                                in_chart_graphic = true;
                            }
                        }
                    }
                    "chart" if in_chart_graphic => {
                        // c:chart/@r:id references the chart part via the slide rels.
                        for attr in e.attributes().flatten() {
                            let key = std::str::from_utf8(attr.key.as_ref()).unwrap_or("");
                            if key == "r:id" || key.ends_with(":id") {
                                chart_r_id =
                                    Some(String::from_utf8_lossy(&attr.value).to_string());
                            }
                        }
                    }
                    "gridCol" if in_table => {
                        // a:tblGrid/a:gridCol/@w in EMU; 12700 EMU = 1pt
                        if let Some(w) = get_attr(&e, "w") {
                            if let Ok(v) = w.parse::<f32>() {
                                tbl_col_widths.push(v / 12700.0);
                            }
                        }
                    }
                    "tr" if in_table => {
                        in_tbl_row = true;
                        tbl_cur_row.clear();
                        // a:tr/@h in EMU; 12700 EMU = 1pt
                        if let Some(h) = get_attr(&e, "h") {
                            if let Ok(v) = h.parse::<f32>() {
                                tbl_row_heights.push(v / 12700.0);
                            }
                        }
                    }
                    "tc" if in_tbl_row => {
                        in_tbl_cell = true;
                        tbl_cur_cell_paragraphs.clear();
                    }
                    "tcPr" if in_tbl_cell => {
                        // Cell properties (fill, borders, spans) — parsed later.
                    }
                    "xfrm" if in_shape => {
                        // a:xfrm/@rot is in 60000ths of a degree; 60000 = 1 degree.
                        shape_has_xfrm = true;
                        if let Some(rot) = get_attr(&e, "rot") {
                            if let Ok(v) = rot.parse::<f32>() {
                                shape_rotation = v / 60000.0;
                            }
                        }
                    }
                    "prstGeom" if in_shape => {
                        // a:prstGeom/@prst = preset shape type, e.g. "rect", "ellipse",
                        // "roundRect", "chevron".
                        if let Some(prst) = get_attr(&e, "prst") {
                            shape_prst = Some(prst);
                        }
                    }
                    "ph" if in_shape => {
                        // p:ph in nvPr — placeholder identity. Missing type defaults
                        // to "obj" (body placeholder) per the schema.
                        shape_ph_type = match get_attr(&e, "type") {
                            Some(t) if !t.is_empty() => Some(t),
                            _ => Some("obj".to_string()),
                        };
                        shape_ph_idx = get_attr(&e, "idx");
                    }
                    "spPr" if in_shape => {
                        in_sp_pr = true;
                    }
                    "ln" if in_sp_pr => {
                        in_ln = true;
                        // Width attribute in EMU; 12700 EMU = 1pt
                        if let Some(w) = get_attr(&e, "w") {
                            if let Ok(v) = w.parse::<f32>() {
                                shape_border_width = Some(v / 12700.0);
                            }
                        }
                    }
                    "off" if in_shape => {
                        if let Some(x) = get_attr(&e, "x") {
                            shape_x = x.parse::<f32>().map(emu_to_pt).unwrap_or(0.0);
                        }
                        if let Some(y) = get_attr(&e, "y") {
                            shape_y = y.parse::<f32>().map(emu_to_pt).unwrap_or(0.0);
                        }
                    }
                    "ext" if in_shape => {
                        if let Some(cx) = get_attr(&e, "cx") {
                            shape_w = cx.parse::<f32>().map(emu_to_pt).unwrap_or(0.0);
                        }
                        if let Some(cy) = get_attr(&e, "cy") {
                            shape_h = cy.parse::<f32>().map(emu_to_pt).unwrap_or(0.0);
                        }
                    }
                    "blipFill" if in_shape => {
                        shape_is_image = true;
                    }
                    "blip" if in_shape && shape_is_image => {
                        // r:embed attribute for image reference
                        for attr in e.attributes().flatten() {
                            let key = std::str::from_utf8(attr.key.as_ref()).unwrap_or("");
                            if key == "r:embed" || key.ends_with(":embed") {
                                shape_image_r_id =
                                    Some(String::from_utf8_lossy(&attr.value).to_string());
                            }
                        }
                    }
                    "p" if in_shape => {
                        in_paragraph = true;
                        para_runs.clear();
                        para_alignment = None;
                        para_line_spacing = None;
                        para_space_before = None;
                        para_space_after = None;
                        in_ln_spc = false;
                        in_spc_bef = false;
                        in_spc_aft = false;
                        para_lvl = 0;
                        para_mar_l = None;
                        para_indent = None;
                        para_bullet = SlideBullet::Inherit;
                        para_bullet_font = None;
                    }
                    "bodyPr" if in_shape => {
                        in_body_pr = true;
                        if let Some(v) = get_attr(&e, "lIns") {
                            if let Ok(v) = v.parse::<f32>() {
                                shape_l_ins = emu_to_pt(v);
                            }
                        }
                        if let Some(v) = get_attr(&e, "rIns") {
                            if let Ok(v) = v.parse::<f32>() {
                                shape_r_ins = emu_to_pt(v);
                            }
                        }
                        if let Some(v) = get_attr(&e, "tIns") {
                            if let Ok(v) = v.parse::<f32>() {
                                shape_t_ins = emu_to_pt(v);
                            }
                        }
                        if let Some(v) = get_attr(&e, "bIns") {
                            if let Ok(v) = v.parse::<f32>() {
                                shape_b_ins = emu_to_pt(v);
                            }
                        }
                        // Spec #6: a:bodyPr/@anchor — vertical text anchoring.
                        // "t" (top) is the default; a value here wins over the
                        // placeholder chain.
                        if let Some(a) = get_attr(&e, "anchor") {
                            shape_anchor = Some(a);
                        }
                    }
                    "pPr" if in_paragraph => {
                        // Spec #6: a:pPr/@algn — the paragraph's own alignment
                        // (wins over the master txStyles level at render time).
                        if let Some(algn) = get_attr(&e, "algn") {
                            para_alignment = Some(parse_alignment_attr(&algn));
                        }
                        // Spec #8: outline level + indents (a:pPr/@lvl/@marL/@indent)
                        if let Some(lvl) = get_attr(&e, "lvl") {
                            if let Ok(v) = lvl.parse::<u32>() {
                                para_lvl = v;
                            }
                        }
                        if let Some(mar_l) = get_attr(&e, "marL") {
                            if let Ok(v) = mar_l.parse::<f32>() {
                                para_mar_l = Some(emu_to_pt(v));
                            }
                        }
                        if let Some(indent) = get_attr(&e, "indent") {
                            if let Ok(v) = indent.parse::<f32>() {
                                para_indent = Some(emu_to_pt(v));
                            }
                        }
                    }
                    "lnSpc" if in_paragraph => {
                        in_ln_spc = true;
                    }
                    "spcBef" if in_paragraph => {
                        in_spc_bef = true;
                    }
                    "spcAft" if in_paragraph => {
                        in_spc_aft = true;
                    }
                    // Spec #8 bullet marker children of a:pPr (Start form).
                    "buFont" if in_paragraph => {
                        if let Some(tf) = get_attr(&e, "typeface") {
                            if !tf.is_empty() {
                                para_bullet_font = Some(tf);
                            }
                        }
                    }
                    "buChar" if in_paragraph => {
                        if let Some(ch) = get_attr(&e, "char") {
                            let mut chars = ch.chars();
                            if let Some(c) = chars.next() {
                                para_bullet = SlideBullet::Char {
                                    ch: c,
                                    font: para_bullet_font.take(),
                                };
                            }
                        }
                    }
                    "buNone" if in_paragraph => {
                        para_bullet = SlideBullet::None;
                    }
                    "buAutoNum" if in_paragraph => {
                        let kind = get_attr(&e, "type").unwrap_or_else(|| "arabicPeriod".to_string());
                        let start_at = get_attr(&e, "startAt").and_then(|v| v.parse::<u32>().ok());
                        para_bullet = SlideBullet::AutoNum { kind, start_at };
                    }
                    "r" if in_paragraph => {
                        in_run = true;
                        run_text.clear();
                        run_bold = false;
                        run_italic = false;
                        run_font_size = None;
                        run_color = None;
                        run_font_family = None;
                    }
                    "rPr" if in_run => {
                        if let Some(b) = get_attr(&e, "b") {
                            run_bold = b == "1" || b == "true";
                        }
                        if let Some(i) = get_attr(&e, "i") {
                            run_italic = i == "1" || i == "true";
                        }
                        if let Some(sz) = get_attr(&e, "sz") {
                            // Font size in hundredths of a point
                            if let Ok(v) = sz.parse::<f32>() {
                                run_font_size = Some(v / 100.0);
                            }
                        }
                    }
                    "solidFill" => {} // container — context determines where color goes
                    "srgbClr" => {
                        if let Some(val) = get_attr(&e, "val") {
                            if in_bg_pr {
                                slide_background_color = Some(val);
                            } else if in_ln && in_sp_pr {
                                shape_border_color = Some(val);
                            } else if in_sp_pr && !in_ln {
                                shape_fill_color = Some(val);
                            } else if in_run {
                                run_color = Some(val);
                            }
                        }
                    }
                    "schemeClr" => {
                        if let Some(val) = get_attr(&e, "val") {
                            let hex = theme_colors.get(&val).cloned().unwrap_or_else(|| scheme_color_to_hex(&val));
                            if in_bg_pr {
                                slide_background_color = Some(hex);
                            } else if in_ln && in_sp_pr {
                                shape_border_color = Some(hex);
                            } else if in_sp_pr && !in_ln {
                                shape_fill_color = Some(hex);
                            } else if in_run {
                                run_color = Some(hex);
                            }
                        }
                    }
                    "latin" | "ea" | "cs" if in_run => {
                        if let Some(typeface) = get_attr(&e, "typeface") {
                            if run_font_family.is_none() {
                                run_font_family = Some(typeface);
                            }
                        }
                    }
                    "t" if in_run => {
                        in_text = true;
                    }
                    _ => {}
                }
            }
            Event::Empty(e) => {
                let name = local_name(e.name().as_ref());
                match name.as_str() {
                    "gridCol" if in_table => {
                        // a:gridCol is typically self-closing.
                        if let Some(w) = get_attr(&e, "w") {
                            if let Ok(v) = w.parse::<f32>() {
                                tbl_col_widths.push(v / 12700.0);
                            }
                        }
                    }
                    "tr" if in_table => {
                        // Self-closing row (rare) — records height only.
                        if let Some(h) = get_attr(&e, "h") {
                            if let Ok(v) = h.parse::<f32>() {
                                tbl_row_heights.push(v / 12700.0);
                            }
                        }
                    }
                    "tc" if in_tbl_row => {
                        // Self-closing cell — an empty cell.
                        tbl_cur_row.push(TableCell {
                            paragraphs: Vec::new(),
                        });
                    }
                    "ln" if in_sp_pr => {
                        // Empty <a:ln/> — no border content, just attributes
                        if let Some(w) = get_attr(&e, "w") {
                            if let Ok(v) = w.parse::<f32>() {
                                shape_border_width = Some(v / 12700.0);
                            }
                        }
                    }
                    "xfrm" if in_shape => {
                        shape_has_xfrm = true;
                        if let Some(rot) = get_attr(&e, "rot") {
                            if let Ok(v) = rot.parse::<f32>() {
                                shape_rotation = v / 60000.0;
                            }
                        }
                    }
                    "prstGeom" if in_shape => {
                        if let Some(prst) = get_attr(&e, "prst") {
                            shape_prst = Some(prst);
                        }
                    }
                    "ph" if in_shape => {
                        shape_ph_type = match get_attr(&e, "type") {
                            Some(t) if !t.is_empty() => Some(t),
                            _ => Some("obj".to_string()),
                        };
                        shape_ph_idx = get_attr(&e, "idx");
                    }
                    "off" if in_shape => {
                        if let Some(x) = get_attr(&e, "x") {
                            shape_x = x.parse::<f32>().map(emu_to_pt).unwrap_or(0.0);
                        }
                        if let Some(y) = get_attr(&e, "y") {
                            shape_y = y.parse::<f32>().map(emu_to_pt).unwrap_or(0.0);
                        }
                    }
                    "ext" if in_shape => {
                        if let Some(cx) = get_attr(&e, "cx") {
                            shape_w = cx.parse::<f32>().map(emu_to_pt).unwrap_or(0.0);
                        }
                        if let Some(cy) = get_attr(&e, "cy") {
                            shape_h = cy.parse::<f32>().map(emu_to_pt).unwrap_or(0.0);
                        }
                    }
                    "chart" if in_chart_graphic => {
                        // c:chart/@r:id references the chart part via the slide rels.
                        // The element is typically self-closing (<c:chart r:id=".."/>),
                        // so the same extraction as the Start arm must run here.
                        for attr in e.attributes().flatten() {
                            let key = std::str::from_utf8(attr.key.as_ref()).unwrap_or("");
                            if key == "r:id" || key.ends_with(":id") {
                                chart_r_id =
                                    Some(String::from_utf8_lossy(&attr.value).to_string());
                            }
                        }
                    }
                    "blip" if in_shape && shape_is_image => {
                        for attr in e.attributes().flatten() {
                            let key = std::str::from_utf8(attr.key.as_ref()).unwrap_or("");
                            if key == "r:embed" || key.ends_with(":embed") {
                                shape_image_r_id =
                                    Some(String::from_utf8_lossy(&attr.value).to_string());
                            }
                        }
                    }
                    "rPr" if in_run => {
                        if let Some(b) = get_attr(&e, "b") {
                            run_bold = b == "1" || b == "true";
                        }
                        if let Some(i) = get_attr(&e, "i") {
                            run_italic = i == "1" || i == "true";
                        }
                        if let Some(sz) = get_attr(&e, "sz") {
                            if let Ok(v) = sz.parse::<f32>() {
                                run_font_size = Some(v / 100.0);
                            }
                        }
                    }
                    "bodyPr" if in_shape => {
                        // Self-closing <p:bodyPr .../> with inset attributes.
                        if let Some(v) = get_attr(&e, "lIns") {
                            if let Ok(v) = v.parse::<f32>() {
                                shape_l_ins = emu_to_pt(v);
                            }
                        }
                        if let Some(v) = get_attr(&e, "rIns") {
                            if let Ok(v) = v.parse::<f32>() {
                                shape_r_ins = emu_to_pt(v);
                            }
                        }
                        if let Some(v) = get_attr(&e, "tIns") {
                            if let Ok(v) = v.parse::<f32>() {
                                shape_t_ins = emu_to_pt(v);
                            }
                        }
                        if let Some(v) = get_attr(&e, "bIns") {
                            if let Ok(v) = v.parse::<f32>() {
                                shape_b_ins = emu_to_pt(v);
                            }
                        }
                        // Spec #6: a:bodyPr/@anchor — vertical text anchoring.
                        if let Some(a) = get_attr(&e, "anchor") {
                            shape_anchor = Some(a);
                        }
                    }
                    "pPr" if in_paragraph => {
                        if let Some(algn) = get_attr(&e, "algn") {
                            para_alignment = Some(parse_alignment_attr(&algn));
                        }
                        if let Some(lvl) = get_attr(&e, "lvl") {
                            if let Ok(v) = lvl.parse::<u32>() {
                                para_lvl = v;
                            }
                        }
                        if let Some(mar_l) = get_attr(&e, "marL") {
                            if let Ok(v) = mar_l.parse::<f32>() {
                                para_mar_l = Some(emu_to_pt(v));
                            }
                        }
                        if let Some(indent) = get_attr(&e, "indent") {
                            if let Ok(v) = indent.parse::<f32>() {
                                para_indent = Some(emu_to_pt(v));
                            }
                        }
                    }
                    // Spec #8 bullet marker children of a:pPr (self-closing form).
                    "buFont" if in_paragraph => {
                        if let Some(tf) = get_attr(&e, "typeface") {
                            if !tf.is_empty() {
                                para_bullet_font = Some(tf);
                            }
                        }
                    }
                    "buChar" if in_paragraph => {
                        if let Some(ch) = get_attr(&e, "char") {
                            let mut chars = ch.chars();
                            if let Some(c) = chars.next() {
                                para_bullet = SlideBullet::Char {
                                    ch: c,
                                    font: para_bullet_font.take(),
                                };
                            }
                        }
                    }
                    "buNone" if in_paragraph => {
                        para_bullet = SlideBullet::None;
                    }
                    "buAutoNum" if in_paragraph => {
                        let kind = get_attr(&e, "type").unwrap_or_else(|| "arabicPeriod".to_string());
                        let start_at = get_attr(&e, "startAt").and_then(|v| v.parse::<u32>().ok());
                        para_bullet = SlideBullet::AutoNum { kind, start_at };
                    }
                    "spcPct" if in_ln_spc => {
                        // a:lnSpc/a:spcPct/@val is the multiple in 100000ths.
                        if let Some(v) = get_attr(&e, "val") {
                            if let Ok(x) = v.parse::<f32>() {
                                para_line_spacing = Some(x / 100000.0);
                            }
                        }
                    }
                    "spcPts" if in_spc_bef || in_spc_aft => {
                        // a:spcBef/spcAft/a:spcPts/@val is in 100ths of a point.
                        if let Some(v) = get_attr(&e, "val") {
                            if let Ok(x) = v.parse::<f32>() {
                                let pt = x / 100.0;
                                if in_spc_bef {
                                    para_space_before = Some(pt);
                                } else {
                                    para_space_after = Some(pt);
                                }
                            }
                        }
                    }
                    "srgbClr" => {
                        if let Some(val) = get_attr(&e, "val") {
                            if in_bg_pr {
                                slide_background_color = Some(val);
                            } else if in_ln && in_sp_pr {
                                shape_border_color = Some(val);
                            } else if in_sp_pr && !in_ln {
                                shape_fill_color = Some(val);
                            } else if in_run {
                                run_color = Some(val);
                            }
                        }
                    }
                    "schemeClr" => {
                        if let Some(val) = get_attr(&e, "val") {
                            let hex = theme_colors.get(&val).cloned().unwrap_or_else(|| scheme_color_to_hex(&val));
                            if in_bg_pr {
                                slide_background_color = Some(hex);
                            } else if in_ln && in_sp_pr {
                                shape_border_color = Some(hex);
                            } else if in_sp_pr && !in_ln {
                                shape_fill_color = Some(hex);
                            } else if in_run {
                                run_color = Some(hex);
                            }
                        }
                    }
                    "latin" | "ea" | "cs" if in_run => {
                        if let Some(typeface) = get_attr(&e, "typeface") {
                            if run_font_family.is_none() {
                                run_font_family = Some(typeface);
                            }
                        }
                    }
                    _ => {}
                }
            }
            Event::End(e) => {
                let name = local_name(e.name().as_ref());
                _depth -= 1;

                match name.as_str() {
                    "bg" => {
                        in_bg = false;
                    }
                    "bgPr" => {
                        in_bg_pr = false;
                    }
                    "spTree" => {
                        in_sp_tree = false;
                    }
                    "spPr" if in_sp_pr => {
                        in_sp_pr = false;
                    }
                    "ln" if in_ln => {
                        in_ln = false;
                    }
                    "sp" | "pic" if in_shape => {
                        let content = if shape_is_image {
                            if let Some(ref r_id) = shape_image_r_id {
                                if let Some(rel) = rels.get(r_id) {
                                    // Load image data from archive
                                    let image_path =
                                        resolve_slide_relative_path(slide_rels_path, &rel.target);
                                    let data = archive
                                        .read_binary_part(&image_path)
                                        .unwrap_or_default();
                                    let ct = detect_content_type(&rel.target);
                                    ShapeContent::Image {
                                        data,
                                        content_type: ct,
                                    }
                                } else {
                                    ShapeContent::Placeholder
                                }
                            } else {
                                ShapeContent::Placeholder
                            }
                        } else if shape_prst.is_some() {
                            // Preset-geometry AutoShape; may or may not carry text.
                            ShapeContent::AutoShape {
                                paragraphs: std::mem::take(&mut shape_paragraphs),
                            }
                        } else if !shape_paragraphs.is_empty() {
                            ShapeContent::TextBox {
                                paragraphs: std::mem::take(&mut shape_paragraphs),
                            }
                        } else {
                            ShapeContent::Placeholder
                        };

                        // Spec #3: a placeholder without an explicit xfrm inherits
                        // the geometry of the matching slideLayout placeholder,
                        // falling back to the slideMaster placeholder geometry.
                        // The ph-key match normalizes type: a slide body placeholder
                        // is written `<p:ph idx="1"/>` (type -> "obj") while the
                        // master writes `<p:ph type="body" idx="1"/>`; "obj" and
                        // "body" denote the same placeholder slot.
                        let (use_x, use_y, use_w, use_h) = if !shape_has_xfrm {
                            if let Some((gx, gy, gw, gh)) = lookup_ph_geom(
                                &layout_ph_geoms,
                                master_ph_geoms,
                                shape_ph_type.as_ref(),
                                shape_ph_idx.as_ref(),
                            ) {
                                (gx, gy, gw, gh)
                            } else {
                                (shape_x, shape_y, shape_w, shape_h)
                            }
                        } else {
                            (shape_x, shape_y, shape_w, shape_h)
                        };

                        // Spec #6: a:bodyPr/@anchor resolved through the placeholder
                        // chain (slide -> layout -> master). A direct anchor on the
                        // shape wins; otherwise a placeholder inherits the anchor
                        // of the matching layout/master placeholder (empty bodyPr
                        // in the chain = inherit).
                        let resolved_anchor = match shape_anchor.take() {
                            Some(a) => Some(a),
                            None => {
                                if shape_ph_type.is_some() {
                                    lookup_ph_anchor(
                                        &layout_ph_anchors,
                                        master_ph_anchors,
                                        shape_ph_type.as_ref(),
                                        shape_ph_idx.as_ref(),
                                    )
                                } else {
                                    None
                                }
                            }
                        };

                        shapes.push(Shape {
                            x: use_x,
                            y: use_y,
                            width: use_w,
                            height: use_h,
                            rotation: shape_rotation,
                            shape_type: shape_prst.take(),
                            ph_type: shape_ph_type.take(),
                            content,
                            fill_color: shape_fill_color.take(),
                            border_color: shape_border_color.take(),
                            border_width: shape_border_width.take(),
                            l_ins: shape_l_ins,
                            r_ins: shape_r_ins,
                            t_ins: shape_t_ins,
                            b_ins: shape_b_ins,
                            anchor: resolved_anchor,
                        });
                        in_shape = false;
                    }
                    "p" if in_paragraph => {
                        in_paragraph = false;
                        let para = SlideParagraph {
                            runs: std::mem::take(&mut para_runs),
                            alignment: para_alignment,
                            line_spacing: para_line_spacing,
                            space_before: para_space_before,
                            space_after: para_space_after,
                            lvl: para_lvl,
                            mar_l: para_mar_l,
                            indent: para_indent,
                            bullet: std::mem::take(&mut para_bullet),
                        };
                        if in_tbl_cell {
                            tbl_cur_cell_paragraphs.push(para);
                        } else {
                            shape_paragraphs.push(para);
                        }
                    }
                    "bodyPr" if in_body_pr => {
                        in_body_pr = false;
                    }
                    "lnSpc" if in_ln_spc => {
                        in_ln_spc = false;
                    }
                    "spcBef" if in_spc_bef => {
                        in_spc_bef = false;
                    }
                    "spcAft" if in_spc_aft => {
                        in_spc_aft = false;
                    }
                    "tc" if in_tbl_cell => {
                        in_tbl_cell = false;
                        tbl_cur_row.push(TableCell {
                            paragraphs: std::mem::take(&mut tbl_cur_cell_paragraphs),
                        });
                    }
                    "tr" if in_tbl_row => {
                        in_tbl_row = false;
                        tbl_rows.push(std::mem::take(&mut tbl_cur_row));
                    }
                    "tbl" if in_table => {
                        in_table = false;
                    }
                    "graphicFrame" if in_graphic_frame => {
                        in_graphic_frame = false;
                        in_shape = false;
                        let content = if in_chart_graphic {
                            // A chart graphicFrame. Resolve the chart part via the
                            // slide rels and parse it (bar chart: series/categories/
                            // values from the cached data).
                            if let Some(ref r_id) = chart_r_id {
                                if let Some(rel) = rels.get(r_id) {
                                    let chart_path = resolve_slide_relative_path(
                                        slide_rels_path,
                                        &rel.target,
                                    );
                                    match archive.try_read_part(&chart_path) {
                                        Ok(Some(chart_xml)) => {
                                            parse_chart(&chart_xml).map(|chart| {
                                                ShapeContent::Chart { chart }
                                            }).unwrap_or_else(|_| ShapeContent::Unsupported {
                                                element_type: "chart".to_string(),
                                            })
                                        }
                                        _ => ShapeContent::Unsupported {
                                            element_type: "chart".to_string(),
                                        },
                                    }
                                } else {
                                    ShapeContent::Unsupported {
                                        element_type: "chart".to_string(),
                                    }
                                }
                            } else {
                                ShapeContent::Unsupported {
                                    element_type: "chart".to_string(),
                                }
                            }
                        } else if !tbl_rows.is_empty() {
                            ShapeContent::Table {
                                table: Table {
                                    col_widths: std::mem::take(&mut tbl_col_widths),
                                    row_heights: std::mem::take(&mut tbl_row_heights),
                                    rows: std::mem::take(&mut tbl_rows),
                                },
                            }
                        } else {
                            ShapeContent::Unsupported {
                                element_type: "graphicFrame".to_string(),
                            }
                        };
                        shapes.push(Shape {
                            x: shape_x,
                            y: shape_y,
                            width: shape_w,
                            height: shape_h,
                            rotation: shape_rotation,
                            shape_type: shape_prst.take(),
                            ph_type: None,
                            content,
                            fill_color: shape_fill_color.take(),
                            border_color: shape_border_color.take(),
                            border_width: shape_border_width.take(),
                            l_ins: shape_l_ins,
                            r_ins: shape_r_ins,
                            t_ins: shape_t_ins,
                            b_ins: shape_b_ins,
                            anchor: None,
                        });
                    }
                    "r" if in_run => {
                        in_run = false;
                        if !run_text.is_empty() {
                            para_runs.push(SlideRun {
                                text: std::mem::take(&mut run_text),
                                font_size: run_font_size,
                                bold: run_bold,
                                italic: run_italic,
                                color: run_color.take(),
                                font_family: run_font_family.take(),
                            });
                        }
                    }
                    "t" => {
                        in_text = false;
                    }
                    _ => {}
                }
            }
            Event::Text(e) => {
                if in_text && in_run {
                    let text = e.unescape()?.to_string();
                    run_text.push_str(&text);
                }
            }
            Event::Eof => break,
            _ => {}
        }
    }

    Ok(Slide {
        index: slide_index,
        shapes,
        background_color: slide_background_color,
    })
}

/// Build a (ph_type, ph_idx) -> (x,y,w,h) geometry map from a slideLayout XML.
/// Spec #3: slide placeholders with no explicit xfrm inherit these geometries.
/// Walk each `<p:sp>` in the layout: capture the `p:ph` (type/idx, inside nvPr)
/// and the `a:xfrm/a:off` + `a:ext`; on `sp` end, if a ph was seen AND an xfrm
/// was seen, insert into the map. Values are converted to points.
/// Parse a slideLayout / slideMaster XML into its placeholder geometry map
/// AND its placeholder bodyPr anchor map (Spec #6). An empty/absent `<a:bodyPr>`
/// contributes no anchor (inheritance); a bodyPr with `anchor` contributes it.
/// Returns (geoms, anchors), both keyed by (ph_type, ph_idx).
fn parse_layout_ph_info(
    xml: &str,
) -> Result<
    (
        HashMap<(Option<String>, Option<String>), (f32, f32, f32, f32)>,
        HashMap<(Option<String>, Option<String>), String>,
    ),
    PptxError,
> {
    let mut geoms: HashMap<(Option<String>, Option<String>), (f32, f32, f32, f32)> = HashMap::new();
    let mut anchors: HashMap<(Option<String>, Option<String>), String> = HashMap::new();
    let mut reader = Reader::from_str(xml);
    let mut in_sp_tree = false;
    let mut in_sp = false;
    let mut in_xfrm = false;
    let mut in_body_pr = false;
    let mut ph_type: Option<String> = None;
    let mut ph_idx: Option<String> = None;
    let mut shape_anchor: Option<String> = None;
    let mut xfrm_seen = false;
    let mut x: f32 = 0.0;
    let mut y: f32 = 0.0;
    let mut w: f32 = 0.0;
    let mut h: f32 = 0.0;

    loop {
        match reader.read_event()? {
            Event::Start(e) => {
                let name = local_name(e.name().as_ref());
                match name.as_str() {
                    "spTree" => {
                        in_sp_tree = true;
                    }
                    "sp" if in_sp_tree => {
                        in_sp = true;
                        ph_type = None;
                        ph_idx = None;
                        shape_anchor = None;
                        xfrm_seen = false;
                        in_body_pr = false;
                        x = 0.0;
                        y = 0.0;
                        w = 0.0;
                        h = 0.0;
                    }
                    "ph" if in_sp => {
                        ph_type = match get_attr(&e, "type") {
                            Some(t) if !t.is_empty() => Some(t),
                            _ => Some("obj".to_string()),
                        };
                        ph_idx = get_attr(&e, "idx");
                    }
                    "bodyPr" if in_sp => {
                        in_body_pr = true;
                        if let Some(a) = get_attr(&e, "anchor") {
                            shape_anchor = Some(a);
                        }
                    }
                    "xfrm" if in_sp => {
                        in_xfrm = true;
                        xfrm_seen = true;
                    }
                    "off" if in_xfrm => {
                        if let Some(v) = get_attr(&e, "x") {
                            if let Ok(v) = v.parse::<f32>() {
                                x = emu_to_pt(v);
                            }
                        }
                        if let Some(v) = get_attr(&e, "y") {
                            if let Ok(v) = v.parse::<f32>() {
                                y = emu_to_pt(v);
                            }
                        }
                    }
                    "ext" if in_xfrm => {
                        if let Some(v) = get_attr(&e, "cx") {
                            if let Ok(v) = v.parse::<f32>() {
                                w = emu_to_pt(v);
                            }
                        }
                        if let Some(v) = get_attr(&e, "cy") {
                            if let Ok(v) = v.parse::<f32>() {
                                h = emu_to_pt(v);
                            }
                        }
                    }
                    _ => {}
                }
            }
            Event::Empty(e) => {
                let name = local_name(e.name().as_ref());
                match name.as_str() {
                    "ph" if in_sp => {
                        ph_type = match get_attr(&e, "type") {
                            Some(t) if !t.is_empty() => Some(t),
                            _ => Some("obj".to_string()),
                        };
                        ph_idx = get_attr(&e, "idx");
                    }
                    "bodyPr" if in_sp => {
                        // Self-closing <a:bodyPr .../> with an anchor attribute.
                        if let Some(a) = get_attr(&e, "anchor") {
                            shape_anchor = Some(a);
                        }
                    }
                    "off" if in_xfrm => {
                        if let Some(v) = get_attr(&e, "x") {
                            if let Ok(v) = v.parse::<f32>() {
                                x = emu_to_pt(v);
                            }
                        }
                        if let Some(v) = get_attr(&e, "y") {
                            if let Ok(v) = v.parse::<f32>() {
                                y = emu_to_pt(v);
                            }
                        }
                    }
                    "ext" if in_xfrm => {
                        if let Some(v) = get_attr(&e, "cx") {
                            if let Ok(v) = v.parse::<f32>() {
                                w = emu_to_pt(v);
                            }
                        }
                        if let Some(v) = get_attr(&e, "cy") {
                            if let Ok(v) = v.parse::<f32>() {
                                h = emu_to_pt(v);
                            }
                        }
                    }
                    _ => {}
                }
            }
            Event::End(e) => {
                let name = local_name(e.name().as_ref());
                match name.as_str() {
                    "bodyPr" if in_body_pr => {
                        in_body_pr = false;
                    }
                    "xfrm" if in_xfrm => {
                        in_xfrm = false;
                    }
                    "sp" if in_sp => {
                        if let Some(pt) = ph_type.as_ref() {
                            if xfrm_seen {
                                geoms.insert(
                                    (Some(pt.clone()), ph_idx.clone()),
                                    (x, y, w, h),
                                );
                            }
                            if let Some(a) = shape_anchor.take() {
                                anchors.insert((Some(pt.clone()), ph_idx.clone()), a);
                            }
                        }
                        in_sp = false;
                    }
                    "spTree" if in_sp_tree => {
                        in_sp_tree = false;
                    }
                    _ => {}
                }
            }
            Event::Eof => break,
            _ => {}
        }
    }

    Ok((geoms, anchors))
}

/// Look up a placeholder geometry with ph-key normalization, checking the
/// slideLayout map first, then the slideMaster map.
///
/// A slide placeholder carries `(ph_type, ph_idx)`. The same placeholder slot
/// can be written differently across slide / layout / master:
///   - slide body:    `<p:ph idx="1"/>`              -> ("obj", Some("1"))
///   - master body:   `<p:ph type="body" idx="1"/>`  -> ("body", Some("1"))
/// So "obj" and "body" are treated as equivalent for a given idx, and an
/// idx-only match is the final fallback.
fn lookup_ph_geom(
    layout: &HashMap<(Option<String>, Option<String>), (f32, f32, f32, f32)>,
    master: &HashMap<(Option<String>, Option<String>), (f32, f32, f32, f32)>,
    ph_type: Option<&String>,
    ph_idx: Option<&String>,
) -> Option<(f32, f32, f32, f32)> {
    // Candidate keys in priority order: exact, then obj/body-equivalence, then
    // idx-only. For a type-only placeholder (title: ("title", None)) the exact
    // key is already the only sensible one.
    let mut keys: Vec<(Option<String>, Option<String>)> = Vec::new();
    keys.push((ph_type.cloned(), ph_idx.cloned()));
    if let Some(idx) = ph_idx {
        let idx = idx.clone();
        keys.push((Some("body".to_string()), Some(idx.clone())));
        keys.push((Some("obj".to_string()), Some(idx.clone())));
        keys.push((None, Some(idx)));
    } else if let Some(ty) = ph_type {
        keys.push((Some(ty.clone()), None));
    }
    for k in &keys {
        if let Some(&g) = layout.get(k) {
            return Some(g);
        }
        if let Some(&g) = master.get(k) {
            return Some(g);
        }
    }
    None
}

/// Spec #6: look up a placeholder's bodyPr anchor through the same
/// (ph_type, ph_idx) key chain as `lookup_ph_geom` — exact key first, then
/// the ctrTitle -> title alias, then obj/body-equivalence, then idx-only.
/// The layout map wins over the master map. Returns None when no placeholder
/// in the chain declares an anchor (= the bodyPr default, "t").
fn lookup_ph_anchor(
    layout: &HashMap<(Option<String>, Option<String>), String>,
    master: &HashMap<(Option<String>, Option<String>), String>,
    ph_type: Option<&String>,
    ph_idx: Option<&String>,
) -> Option<String> {
    let mut keys: Vec<(Option<String>, Option<String>)> = Vec::new();
    keys.push((ph_type.cloned(), ph_idx.cloned()));
    // "ctrTitle" (slide) and "title" (layout/master) are the same slot.
    if let Some(ty) = ph_type {
        if ty == "ctrTitle" {
            keys.push((Some("title".to_string()), ph_idx.cloned()));
        }
    }
    if let Some(idx) = ph_idx {
        let idx = idx.clone();
        keys.push((Some("body".to_string()), Some(idx.clone())));
        keys.push((Some("obj".to_string()), Some(idx.clone())));
        keys.push((None, Some(idx)));
    } else if let Some(ty) = ph_type {
        keys.push((Some(ty.clone()), None));
    }
    for k in &keys {
        if let Some(a) = layout.get(k) {
            return Some(a.clone());
        }
        if let Some(a) = master.get(k) {
            return Some(a.clone());
        }
    }
    None
}

/// Resolve a target path relative to the slide location.
/// e.g., slide rels at "ppt/slides/_rels/slide1.xml.rels", target "../media/image1.png"
///   -> "ppt/media/image1.png"
fn resolve_slide_relative_path(rels_path: &str, target: &str) -> String {
    if target.starts_with('/') {
        return target.trim_start_matches('/').to_string();
    }

    // Get the directory of the slide (parent of _rels)
    // rels_path: "ppt/slides/_rels/slide1.xml.rels"
    // slide dir: "ppt/slides/"
    let slide_dir = if let Some(pos) = rels_path.rfind("/_rels/") {
        &rels_path[..pos + 1] // "ppt/slides/"
    } else if let Some(pos) = rels_path.rfind('/') {
        &rels_path[..pos + 1]
    } else {
        ""
    };

    // Resolve "../" segments
    let mut base_parts: Vec<&str> = slide_dir
        .split('/')
        .filter(|s| !s.is_empty())
        .collect();
    for segment in target.split('/') {
        match segment {
            ".." => {
                // Prevent escaping beyond archive root
                if base_parts.is_empty() {
                    return String::new(); // reject traversal beyond root
                }
                base_parts.pop();
            }
            "." | "" => {}
            s => base_parts.push(s),
        }
    }

    base_parts.join("/")
}

/// Parse a chart part (chartN.xml) into a Chart IR.
///
/// Extracts a bar chart's cached data: the series names, categories and values
/// from the `c:barChart`'s `c:ser` children (strCache / numCache). Chart part
/// discovery / parsing is the chart Ra-loop Step 1 (impact order item 8); the
/// Word-measured plot geometry is consumed at render time (Step 2+).
fn parse_chart(xml: &str) -> Result<Chart, PptxError> {
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(false);

    let mut in_bar_chart = false;
    let mut chart_type: Option<String> = None;
    let mut bar_dir: Option<String> = None;
    let mut grouping: Option<String> = None;
    let mut series: Vec<ChartSeries> = Vec::new();
    let mut categories: Vec<String> = Vec::new();
    let mut has_legend = false;
    let mut auto_title_deleted = false;
    let mut explicit_title: Option<String> = None;
    let mut in_title = false;
    let mut in_title_t = false;
    let mut title_text = String::new();
    let mut marker = false;

    // Data-label (`c:dLbls`) state. `in_dlbls` scopes the numFmt handler
    // (the axes also carry `c:numFmt`, so it must not fire outside dLbls).
    let mut in_dlbls = false;
    let mut has_data_labels = false;
    let mut datalabel_position = String::new();
    let mut number_format = String::new();
    let mut show_val = false;
    let mut show_cat_name = false;
    let mut show_ser_name = false;
    let mut show_percent = false;
    let mut show_legend_key = false;
    let mut show_bubble_size = false;

    // Per-`c:ser` state
    let mut in_ser = false;
    // Which cache we're collecting: "tx" | "cat" | "val" | ""
    let mut ser_target = "";
    let mut ser_name: Option<String> = None;
    let mut ser_values: Vec<f64> = Vec::new();
    let mut in_v = false;
    let mut cur_v = String::new();
    let mut ser_categories: Vec<String> = Vec::new();

    loop {
        match reader.read_event() {
            Ok(Event::Start(e)) => {
                let name = local_name(e.name().as_ref());
                match name.as_str() {
                    "pieChart" => {
                        chart_type = Some("pie".to_string());
                    }
                    "lineChart" => {
                        chart_type = Some("line".to_string());
                    }
                    "barChart" => {
                        in_bar_chart = true;
                    }
                    "ser"
                        if in_bar_chart
                            || chart_type.as_deref() == Some("pie")
                            || chart_type.as_deref() == Some("line") =>
                    {
                        in_ser = true;
                        ser_target = "";
                        ser_name = None;
                        ser_values.clear();
                        ser_categories.clear();
                    }
                    "tx" if in_ser => ser_target = "tx",
                    "cat" if in_ser => ser_target = "cat",
                    "val" if in_ser => ser_target = "val",
                    "v" => {
                        in_v = true;
                        cur_v.clear();
                    }
                    // <c:dLbls> is a real START tag in python-pptx output
                    // (it carries the label child elements), so catch it
                    // here; the matching End resets in_dlbls. Word draws
                    // data labels whenever a <c:dLbls> exists.
                    "dLbls" => {
                        in_dlbls = true;
                        has_data_labels = true;
                    }
                    // <c:legend> is a REAL START tag in python-pptx output
                    // ("<c:legend><c:legendPos .../><c:layout/><c:overlay .../>
                    // </c:legend>"), NOT a self-closing <c:legend/>. Word
                    // draws a legend whenever a <c:legend> element exists
                    // (regardless of position/overlay attrs), so catch BOTH
                    // the Start form here and the self-closing form below.
                    "legend" => has_legend = true,
                    // <c:title> is a REAL START tag carrying the explicit
                    // chart-title text: <c:title><c:tx><c:rich><a:p><a:r>
                    // <a:t>Quarterly Revenue</a:t></a:r></a:p>...</c:title>.
                    // Word draws it (Arial 18pt regular) INSTEAD of the
                    // automatic series-name title — chart_title / chart_title2
                    // render-truth 2026-08-07.
                    "title" => in_title = true,
                    // <a:t> (drawingml) is the text-run element inside
                    // c:title/c:tx/c:rich.
                    "t" if in_title => {
                        in_title_t = true;
                        title_text.clear();
                    }
                    _ => {}
                }
            }
            Ok(Event::Empty(e)) => {
                let name = local_name(e.name().as_ref());
                // A v may be empty; treat as no-op but keep flag semantics.
                match name.as_str() {
                    "v" => {
                        // empty value — nothing to collect
                    }
                    // <c:barDir val="col"/> and <c:grouping val="stacked"/>
                    // are SELF-CLOSING CHILD elements of <c:barChart> (NOT
                    // attributes on the barChart tag) — Event::Empty. The
                    // Word-measured stacked-column spec (chart_stacked probe)
                    // keys the stack vs cluster split on chart.grouping.
                    "barDir" => {
                        if let Some(v) = get_attr(&e, "val") {
                            bar_dir = Some(v);
                        }
                    }
                    "grouping" => {
                        if let Some(v) = get_attr(&e, "val") {
                            grouping = Some(v);
                        }
                    }
                    // python-pptx writes a bare self-closing <c:legend/> to
                    // enable a legend (no overlay/position attrs). Any legend
                    // declaration -> has_legend.
                    "legend" => has_legend = true,
                    // Self-closing <c:autoTitleDeleted val="1"/> — python-pptx
                    // writes it when the auto title is explicitly removed
                    // (chart.has_title=False). When absent (or val=0) Word
                    // draws the automatic series-name title and shifts the
                    // pie circle down. This is an Event::Empty (self-closing),
                    // NOT a Start — the same trap as <c:chart r:id=.../>.
                    "autoTitleDeleted" => {
                        if get_attr(&e, "val").as_deref() == Some("1") {
                            auto_title_deleted = true;
                        }
                    }
                    // <c:marker val="1"/> is a SELF-CLOSING CHILD element of
                    // <c:lineChart> (python-pptx writes it for LINE_MARKERS).
                    // marker=1 -> Word draws a filled accent-colour circle at
                    // each data point. Absent or val="0" -> no markers.
                    // Same Event::Empty trap as barDir/grouping/legend/
                    // autoTitleDeleted.
                    "marker" => {
                        if get_attr(&e, "val").as_deref().map(|v| v != "0").unwrap_or(true) {
                            marker = true;
                        }
                    }
                    // Data-label child elements are all SELF-CLOSING
                    // (Event::Empty) in python-pptx output: <c:dLblPos
                    // val="ctr"/>, <c:numFmt sourceLinked="0"
                    // formatCode="0.0%"/>, <c:showVal val="1"/> etc.
                    // All are gated on in_dlbls so the AXES' <c:numFmt>
                    // (which appears outside any dLbls) is never captured.
                    "dLblPos" if in_dlbls => {
                        if let Some(v) = get_attr(&e, "val") {
                            datalabel_position = v;
                        }
                    }
                    "numFmt" if in_dlbls => {
                        if let Some(fmt) = get_attr(&e, "formatCode") {
                            number_format = fmt;
                        }
                    }
                    "showVal" if in_dlbls => {
                        show_val = get_attr(&e, "val").as_deref() == Some("1");
                    }
                    "showCatName" if in_dlbls => {
                        show_cat_name = get_attr(&e, "val").as_deref() == Some("1");
                    }
                    "showSerName" if in_dlbls => {
                        show_ser_name = get_attr(&e, "val").as_deref() == Some("1");
                    }
                    "showPercent" if in_dlbls => {
                        show_percent = get_attr(&e, "val").as_deref() == Some("1");
                    }
                    "showLegendKey" if in_dlbls => {
                        show_legend_key = get_attr(&e, "val").as_deref() == Some("1");
                    }
                    "showBubbleSize" if in_dlbls => {
                        show_bubble_size = get_attr(&e, "val").as_deref() == Some("1");
                    }
                    _ => {}
                }
            }
            Ok(Event::Text(e)) => {
                if in_v {
                    cur_v.push_str(&e.unescape().unwrap_or_default());
                } else if in_title_t {
                    title_text.push_str(&e.unescape().unwrap_or_default());
                }
            }
            Ok(Event::End(e)) => {
                let name = local_name(e.name().as_ref());
                match name.as_str() {
                    "v" => {
                        if in_v {
                            in_v = false;
                            match ser_target {
                                "tx" => {
                                    if ser_name.is_none() && !cur_v.trim().is_empty() {
                                        ser_name = Some(cur_v.trim().to_string());
                                    }
                                }
                                "cat" => {
                                    if !cur_v.trim().is_empty() {
                                        ser_categories.push(cur_v.trim().to_string());
                                    }
                                }
                                "val" => {
                                    if let Ok(v) = cur_v.trim().parse::<f64>() {
                                        ser_values.push(v);
                                    }
                                }
                                _ => {}
                            }
                        }
                    }
                    "ser" => {
                        if in_ser {
                            in_ser = false;
                            ser_target = "";
                            let name = ser_name.take().unwrap_or_default();
                            if categories.is_empty() {
                                categories = std::mem::take(&mut ser_categories);
                            } else {
                                ser_categories.clear();
                            }
                            series.push(ChartSeries {
                                name,
                                values: std::mem::take(&mut ser_values),
                            });
                        }
                    }
                    "barChart" => {
                        in_bar_chart = false;
                    }
                    "dLbls" => {
                        in_dlbls = false;
                    }
                    "t" if in_title => {
                        in_title_t = false;
                    }
                    "title" => {
                        in_title = false;
                        if !title_text.trim().is_empty() {
                            explicit_title = Some(title_text.trim().to_string());
                        }
                    }
                    _ => {}
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => return Err(PptxError::Xml(e)),
            _ => {}
        }
    }

    Ok(Chart {
        chart_type: chart_type.unwrap_or_else(default_chart_type),
        bar_dir: bar_dir.unwrap_or_else(default_chart_bar_dir),
        grouping: grouping.unwrap_or_else(default_chart_grouping),
        series,
        categories,
        has_legend,
        auto_title_deleted,
        explicit_title,
        marker,
        has_data_labels,
        datalabel_position,
        number_format,
        show_val,
        show_cat_name,
        show_ser_name,
        show_percent,
        show_legend_key,
        show_bubble_size,
    })
}

/// Detect content type from file extension.
fn detect_content_type(path: &str) -> Option<String> {
    let lower = path.to_lowercase();
    if lower.ends_with(".png") {
        Some("image/png".to_string())
    } else if lower.ends_with(".jpg") || lower.ends_with(".jpeg") {
        Some("image/jpeg".to_string())
    } else if lower.ends_with(".gif") {
        Some("image/gif".to_string())
    } else if lower.ends_with(".bmp") {
        Some("image/bmp".to_string())
    } else if lower.ends_with(".svg") {
        Some("image/svg+xml".to_string())
    } else if lower.ends_with(".emf") {
        Some("image/x-emf".to_string())
    } else if lower.ends_with(".wmf") {
        Some("image/x-wmf".to_string())
    } else if lower.ends_with(".tiff") || lower.ends_with(".tif") {
        Some("image/tiff".to_string())
    } else {
        None
    }
}

/// Map common OOXML scheme color names to approximate hex values.
fn scheme_color_to_hex(scheme: &str) -> String {
    match scheme {
        "bg1" | "lt1" => "FFFFFF",
        "bg2" | "lt2" => "E7E6E6",
        "tx1" | "dk1" => "000000",
        "tx2" | "dk2" => "44546A",
        "accent1" => "4472C4",
        "accent2" => "ED7D31",
        "accent3" => "A5A5A5",
        "accent4" => "FFC000",
        "accent5" => "5B9BD5",
        "accent6" => "70AD47",
        "hlink" => "0563C1",
        "folHlink" => "954F72",
        _ => "000000",
    }
    .to_string()
}

/// Parse a .pptx file from raw bytes into a Presentation IR.
pub fn parse_pptx(data: &[u8]) -> Result<Presentation, PptxError> {
    let mut archive = OoxmlArchive::new(data)?;

    // 1. Parse presentation.xml for slide list and slide size
    let pres_xml = archive.read_part("ppt/presentation.xml")?;
    let (slide_infos, slide_width, slide_height) = parse_presentation_slides(&pres_xml)?;

    // 2. Parse presentation relationships
    let rels_xml = archive.read_part("ppt/_rels/presentation.xml.rels")?;
    let rels = parse_relationships(&rels_xml)?;

    // Build rId -> target path map
    let rid_to_path: std::collections::HashMap<String, String> = rels
        .into_iter()
        .map(|(id, rel)| (id, rel.target))
        .collect();

    // 2.5. Parse the theme for minor/major latin typefaces and the clrScheme
    // colour map (Spec #10). Standard path is ppt/theme/theme1.xml; a theme
    // without a latin face or an absent theme part falls back to "Calibri",
    // and the colour map falls back to the built-in table per slot.
    let (minor_font, major_font, theme_colors) = match archive.try_read_part("ppt/theme/theme1.xml")? {
        Some(theme_xml) => parse_theme(&theme_xml)?,
        None => (
            "Calibri".to_string(),
            "Calibri".to_string(),
            HashMap::new(),
        ),
    };

    // 2.7. Parse the slide master text styles (Spec #8): the inherited
    // marL/indent/bullet/spcBef per outline level. Standard path is
    // ppt/slideMasters/slideMaster1.xml; absent part -> empty styles.
    let master_styles = match archive.try_read_part("ppt/slideMasters/slideMaster1.xml")? {
        Some(master_xml) => parse_master_txstyles(&master_xml)?,
        None => MasterTxStyles::default(),
    };

    // 2.8. Parse the slide master placeholder geometries AND bodyPr anchors
    // (Spec #8 follow-up + Spec #6). A layout placeholder without an explicit
    // xfrm inherits the master's ph geometry; the slide->layout->master chain
    // resolves in parse_slide. Same (ph_type, ph_idx) map shape as the layout
    // map; the master anchors feed the anchor resolution chain the same way.
    let (master_ph_geoms, master_ph_anchors): (
        HashMap<(Option<String>, Option<String>), (f32, f32, f32, f32)>,
        HashMap<(Option<String>, Option<String>), String>,
    ) = match archive.try_read_part("ppt/slideMasters/slideMaster1.xml")? {
        Some(master_xml) => parse_layout_ph_info(&master_xml).unwrap_or_default(),
        None => Default::default(),
    };

    // 3. Parse each slide
    let mut slides = Vec::new();
    for (i, info) in slide_infos.iter().enumerate() {
        let slide_target = match rid_to_path.get(&info.r_id) {
            Some(target) => {
                if target.starts_with('/') {
                    target.trim_start_matches('/').to_string()
                } else {
                    format!("ppt/{}", target)
                }
            }
            None => {
                log::warn!("No relationship found for slide rId={}, skipping", info.r_id);
                continue;
            }
        };

        // Slide rels path: e.g., "ppt/slides/_rels/slide1.xml.rels"
        let slide_rels_path = {
            if let Some(pos) = slide_target.rfind('/') {
                let dir = &slide_target[..pos + 1];
                let filename = &slide_target[pos + 1..];
                format!("{}/_rels/{}.rels", dir.trim_end_matches('/'), filename)
            } else {
                format!("_rels/{}.rels", slide_target)
            }
        };

        match archive.try_read_part(&slide_target)? {
            Some(slide_xml) => {
                let slide = parse_slide(
                    &slide_xml,
                    i + 1,
                    &mut archive,
                    &slide_rels_path,
                    &master_ph_geoms,
                    &master_ph_anchors,
                    &theme_colors,
                )?;
                slides.push(slide);
            }
            None => {
                log::warn!("Slide file '{}' not found in archive, skipping", slide_target);
            }
        }
    }

    Ok(Presentation {
        slides,
        slide_width,
        slide_height,
        minor_font,
        major_font,
        theme_colors,
        master_styles,
    })
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_parse_presentation_slides() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:sldIdLst>
    <p:sldId id="256" r:id="rId2"/>
    <p:sldId id="257" r:id="rId3"/>
  </p:sldIdLst>
  <p:sldSz cx="9144000" cy="6858000"/>
</p:presentation>"#;
        let (slides, w, h) = parse_presentation_slides(xml).unwrap();
        assert_eq!(slides.len(), 2);
        assert_eq!(slides[0].r_id, "rId2");
        assert_eq!(slides[1].r_id, "rId3");
        // 9144000 EMU / 12700 = 720pt
        assert!((w - 720.0).abs() < 0.1);
        // 6858000 EMU / 12700 = 540pt
        assert!((h - 540.0).abs() < 0.1);
    }

    #[test]
    fn test_resolve_slide_relative_path() {
        assert_eq!(
            resolve_slide_relative_path("ppt/slides/_rels/slide1.xml.rels", "../media/image1.png"),
            "ppt/media/image1.png"
        );
        assert_eq!(
            resolve_slide_relative_path("ppt/slides/_rels/slide1.xml.rels", "image1.png"),
            "ppt/slides/image1.png"
        );
        assert_eq!(
            resolve_slide_relative_path("ppt/slides/_rels/slide1.xml.rels", "/ppt/media/img.png"),
            "ppt/media/img.png"
        );
    }

    #[test]
    fn test_detect_content_type() {
        assert_eq!(detect_content_type("image1.png"), Some("image/png".to_string()));
        assert_eq!(detect_content_type("photo.JPEG"), Some("image/jpeg".to_string()));
        assert_eq!(detect_content_type("logo.svg"), Some("image/svg+xml".to_string()));
        assert_eq!(detect_content_type("unknown.xyz"), None);
    }
}
