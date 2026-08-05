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
    Presentation, Shape, ShapeContent, Slide, SlideAlignment, SlideParagraph, SlideRun, Table,
    TableCell,
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
fn parse_theme(xml: &str) -> Result<(String, String), PptxError> {
    let mut reader = Reader::from_str(xml);
    let mut minor = "Calibri".to_string();
    let mut major = "Calibri".to_string();
    let mut in_minor = false;
    let mut in_major = false;

    loop {
        match reader.read_event()? {
            Event::Start(e) | Event::Empty(e) => {
                let name = local_name(e.name().as_ref());
                match name.as_str() {
                    "minorFont" => in_minor = true,
                    "majorFont" => in_major = true,
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
                    _ => {}
                }
            }
            Event::End(e) => {
                let name = local_name(e.name().as_ref());
                match name.as_str() {
                    "minorFont" => in_minor = false,
                    "majorFont" => in_major = false,
                    _ => {}
                }
            }
            Event::Eof => break,
            _ => {}
        }
    }

    Ok((minor, major))
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

/// Parse a single slide XML into shapes.
fn parse_slide(
    xml: &str,
    slide_index: usize,
    archive: &mut OoxmlArchive,
    slide_rels_path: &str,
) -> Result<Slide, PptxError> {
    // Parse slide relationships for image resolution
    let rels = if let Ok(Some(rels_xml)) = archive.try_read_part(slide_rels_path) {
        parse_relationships(&rels_xml).unwrap_or_default()
    } else {
        Default::default()
    };

    // Spec #3: build a (ph_type, ph_idx) -> (x,y,w,h) geometry map from the
    // referenced slideLayout's placeholders. A slide placeholder with NO explicit
    // xfrm in its spPr inherits the layout placeholder's geometry.
    let layout_ph_geoms: HashMap<(Option<String>, Option<String>), (f32, f32, f32, f32)> = {
        let mut map = HashMap::new();
        for rel in rels.values() {
            if rel.rel_type.ends_with("/slideLayout") {
                let layout_path =
                    resolve_slide_relative_path(slide_rels_path, &rel.target);
                if let Ok(Some(layout_xml)) = archive.try_read_part(&layout_path) {
                    map = parse_layout_ph_geoms(&layout_xml).unwrap_or_default();
                }
                break;
            }
        }
        map
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

    // Shape property context tracking
    let mut in_sp_pr = false; // inside <p:spPr> or <xdr:spPr>
    let mut in_ln = false;    // inside <a:ln> (line/border properties)

    // Paragraph state
    let mut in_paragraph = false;
    let mut para_runs: Vec<SlideRun> = Vec::new();
    let mut para_alignment = SlideAlignment::default();
    // Spec #4: paragraph spacing (a:pPr/a:lnSpc, a:spcBef, a:spcAft)
    let mut para_line_spacing: Option<f32> = None;
    let mut para_space_before: Option<f32> = None;
    let mut para_space_after: Option<f32> = None;
    let mut in_ln_spc = false;
    let mut in_spc_bef = false;
    let mut in_spc_aft = false;

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
                        shape_ph_type = None;
                        shape_ph_idx = None;
                        shape_has_xfrm = false;
                    }
                    "tbl" if in_graphic_frame => {
                        in_table = true;
                        tbl_col_widths.clear();
                        tbl_row_heights.clear();
                        tbl_rows.clear();
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
                        para_alignment = SlideAlignment::default();
                        para_line_spacing = None;
                        para_space_before = None;
                        para_space_after = None;
                        in_ln_spc = false;
                        in_spc_bef = false;
                        in_spc_aft = false;
                    }
                    "pPr" if in_paragraph => {
                        if let Some(algn) = get_attr(&e, "algn") {
                            para_alignment = match algn.as_str() {
                                "ctr" => SlideAlignment::Center,
                                "r" => SlideAlignment::Right,
                                "just" => SlideAlignment::Justify,
                                _ => SlideAlignment::Left,
                            };
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
                            let hex = scheme_color_to_hex(&val);
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
                    "pPr" if in_paragraph => {
                        if let Some(algn) = get_attr(&e, "algn") {
                            para_alignment = match algn.as_str() {
                                "ctr" => SlideAlignment::Center,
                                "r" => SlideAlignment::Right,
                                "just" => SlideAlignment::Justify,
                                _ => SlideAlignment::Left,
                            };
                        }
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
                            let hex = scheme_color_to_hex(&val);
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
                        // the geometry of the matching slideLayout placeholder.
                        let (use_x, use_y, use_w, use_h) = if !shape_has_xfrm {
                            if let Some(ph_type) = shape_ph_type.as_ref() {
                                if let Some(&(lx, ly, lw, lh)) = layout_ph_geoms.get(&(
                                    Some(ph_type.clone()),
                                    shape_ph_idx.clone(),
                                )) {
                                    (lx, ly, lw, lh)
                                } else {
                                    (shape_x, shape_y, shape_w, shape_h)
                                }
                            } else {
                                (shape_x, shape_y, shape_w, shape_h)
                            }
                        } else {
                            (shape_x, shape_y, shape_w, shape_h)
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
                        };
                        if in_tbl_cell {
                            tbl_cur_cell_paragraphs.push(para);
                        } else {
                            shape_paragraphs.push(para);
                        }
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
                        let content = if !tbl_rows.is_empty() {
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
fn parse_layout_ph_geoms(
    xml: &str,
) -> Result<HashMap<(Option<String>, Option<String>), (f32, f32, f32, f32)>, PptxError> {
    let mut map: HashMap<(Option<String>, Option<String>), (f32, f32, f32, f32)> = HashMap::new();
    let mut reader = Reader::from_str(xml);
    let mut in_sp_tree = false;
    let mut in_sp = false;
    let mut in_xfrm = false;
    let mut ph_type: Option<String> = None;
    let mut ph_idx: Option<String> = None;
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
                        xfrm_seen = false;
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
                    "xfrm" if in_xfrm => {
                        in_xfrm = false;
                    }
                    "sp" if in_sp => {
                        if let Some(pt) = ph_type.as_ref() {
                            if xfrm_seen {
                                map.insert(
                                    (Some(pt.clone()), ph_idx.clone()),
                                    (x, y, w, h),
                                );
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

    Ok(map)
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

    // 2.5. Parse the theme for minor/major latin typefaces.
    // Standard path is ppt/theme/theme1.xml; a theme without a latin face or
    // an absent theme part falls back to "Calibri".
    let (minor_font, major_font) = match archive.try_read_part("ppt/theme/theme1.xml")? {
        Some(theme_xml) => parse_theme(&theme_xml)?,
        None => ("Calibri".to_string(), "Calibri".to_string()),
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
                let slide =
                    parse_slide(&slide_xml, i + 1, &mut archive, &slide_rels_path)?;
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
