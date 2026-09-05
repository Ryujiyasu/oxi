// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

use std::collections::HashMap;

use quick_xml::events::Event;
use quick_xml::reader::Reader;
use thiserror::Error;

use oxidocs_common::archive::OoxmlArchive;
use oxidocs_common::relationships::{parse_relationships, Relationship};
use oxidocs_common::xml_utils::{emu_to_pt, get_attr, local_name};

use crate::ir::{
    default_b_ins, default_chart_bar_dir, default_chart_bubble_scale, default_chart_grouping,
    default_chart_hole_size, default_chart_size_represents,
    default_chart_updown_gap,
    default_chart_type, default_l_ins, default_r_ins, default_t_ins, Chart, ChartSeries,
    CellBorder, CustomGeometry, EmbeddedFont, GeomCmd, GeomPath, LineEnd,
    MasterStyleLevel, MasterTxStyles, Presentation, Shape, ShapeContent, Slide,
    SlideAlignment, SlideBackgroundImage, SlideBullet, SlideGradient, SlideGradientStop,
    SlideParagraph,
    SlideRun, Table, TableCell,
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
    let mut in_def_rpr_lvl = false;
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
                    "latin" if in_level && in_def_rpr_lvl => {
                        if let Some(t) = get_attr(&e, "typeface") {
                            if !t.is_empty() && cur_level.font_family.is_none() {
                                cur_level.font_family = Some(t);
                            }
                        }
                    }
                    "defRPr" if in_level => {
                        in_def_rpr_lvl = true;
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
                    "defRPr" if in_def_rpr_lvl => in_def_rpr_lvl = false,
                    n if is_master_lvl(n) && in_level => {
                        in_def_rpr_lvl = false;
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

/// An `a:headEnd` / `a:tailEnd` element. `type="none"` and a missing `@type`
/// both mean "no decoration", which is `None` rather than a stored token --
/// the corpus states `type="none"` explicitly on every table-cell border, so
/// keeping those would triple the field's occupancy for nothing.
fn parse_line_end(e: &quick_xml::events::BytesStart) -> Option<LineEnd> {
    let kind = get_attr(e, "type").filter(|t| t != "none")?;
    Some(LineEnd {
        kind,
        w: get_attr(e, "w").unwrap_or_else(|| "med".to_string()),
        len: get_attr(e, "len").unwrap_or_else(|| "med".to_string()),
    })
}

/// Follow one relationship of `rel_type` out of `rels` and return the target
/// path plus the `_rels` path of that target, so the walk can continue.
fn follow_rel(
    rels: &HashMap<String, Relationship>,
    from_rels_path: &str,
    rel_type: &str,
) -> Option<(String, String)> {
    let rel = rels.values().find(|r| r.rel_type.ends_with(rel_type))?;
    let path = resolve_slide_relative_path(from_rels_path, &rel.target);
    let cut = path.rfind('/').map(|i| i + 1).unwrap_or(0);
    let path_rels = format!("{}_rels/{}.rels", &path[..cut], &path[cut..]);
    Some((path, path_rels))
}

/// The clrScheme a slide's own `schemeClr` values resolve against: the theme of
/// the slideMaster its slideLayout points at, reached by relationship
/// (slide -> layout -> master -> theme).  `None` when any link is missing, so
/// the caller keeps the deck-level map.
fn resolve_slide_theme_colors(
    rels: &HashMap<String, Relationship>,
    slide_rels_path: &str,
    archive: &mut OoxmlArchive,
) -> Option<HashMap<String, String>> {
    let (_, layout_rels_path) = follow_rel(rels, slide_rels_path, "/slideLayout")?;
    let layout_rels =
        parse_relationships(&archive.try_read_part(&layout_rels_path).ok()??).ok()?;
    let (_, master_rels_path) = follow_rel(&layout_rels, &layout_rels_path, "/slideMaster")?;
    let master_rels =
        parse_relationships(&archive.try_read_part(&master_rels_path).ok()??).ok()?;
    let (theme_path, _) = follow_rel(&master_rels, &master_rels_path, "/theme")?;
    let theme_xml = archive.try_read_part(&theme_path).ok()??;
    let (_, _, colors) = parse_theme(&theme_xml).ok()?;
    Some(colors)
}

/// Hand the accumulated `a:custGeom` paths to the shape being closed and reset
/// the accumulator for the next one.
///
/// Returns None for a shape that had no custGeom, for one whose geometry used a
/// command outside the modelled vocabulary, and for one whose paths carry no
/// drawable segment -- in all three the consumer keeps its previous (bounding
/// box) rendering rather than drawing a partial outline.
fn take_custom_geometry(paths: &mut Vec<GeomPath>, unsupported: &mut bool) -> Option<CustomGeometry> {
    let bad = std::mem::replace(unsupported, false);
    let paths = std::mem::take(paths);
    if bad || paths.iter().all(|p| p.commands.is_empty()) {
        return None;
    }
    Some(CustomGeometry {
        paths,
        unsupported: false,
    })
}

/// The per-level text styles each LAYOUT placeholder declares in its own
/// `a:lstStyle`, keyed like the geometry map.
///
/// PowerPoint render-truth (d24 slide 1, 2026-08-18): the deck's master
/// `p:titleStyle` lvl1 carries no `defRPr` size or colour at all, while the
/// layout's ctrTitle placeholder carries
/// `<a:defRPr sz="6000"><a:solidFill><a:schemeClr val="lt1"/>`, and
/// PowerPoint draws the title at 60pt in white (#FFFFFF, measured from its own
/// PDF). Without this the title fell back to the engine's 18pt black.
///
/// This is the placeholder's `a:lstStyle`, NOT the layout's `p:txStyles` --
/// the phfs probe showed PowerPoint ignores the latter.
fn parse_layout_ph_lststyles(
    xml: &str,
    theme_colors: &HashMap<String, String>,
) -> HashMap<(Option<String>, Option<String>), Vec<MasterStyleLevel>> {
    let mut out: HashMap<(Option<String>, Option<String>), Vec<MasterStyleLevel>> = HashMap::new();
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();
    let mut in_sp = false;
    let mut in_lst = false;
    let mut ph_type: Option<String> = None;
    let mut ph_idx: Option<String> = None;
    let mut levels: Vec<MasterStyleLevel> = Vec::new();
    let mut cur_lvl: Option<usize> = None;
    let mut in_def_rpr = false;
    let s_lvlitalic = std::env::var("OXI_LVLITALIC_DISABLE").is_err();
    let s_lvlbold = std::env::var("OXI_LVLBOLD_DISABLE").is_err();
    // `a:highlight` inside a level's defRPr holds a colour element shaped just
    // like the level's own solidFill, so without this flag it is read as the
    // level's TEXT colour -- the same trap the run-level highlight sprang.
    let mut in_lvl_highlight = false;
    let s_highlight_lvl = std::env::var("OXI_HIGHLIGHTLVL_DISABLE").is_err();
    let mut in_ln_spc_lvl = false;
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                let name = local_name(e.name().as_ref());
                match name.as_str() {
                    "sp" => {
                        in_sp = true;
                        ph_type = None;
                        ph_idx = None;
                        levels.clear();
                    }
                    "ph" if in_sp => {
                        ph_type = match get_attr(&e, "type") {
                            Some(t) if !t.is_empty() => Some(t),
                            _ => Some("obj".to_string()),
                        };
                        ph_idx = get_attr(&e, "idx");
                    }
                    "lstStyle" if in_sp => in_lst = true,
                    _ if in_lst && is_master_lvl(&name) => {
                        let idx = name
                            .trim_start_matches("lvl")
                            .trim_end_matches("pPr")
                            .parse::<usize>()
                            .unwrap_or(1)
                            .saturating_sub(1);
                        while levels.len() <= idx {
                            levels.push(MasterStyleLevel::default());
                        }
                        cur_lvl = Some(idx);
                        if let Some(a) = get_attr(&e, "algn") {
                            levels[idx].algn = Some(parse_alignment_attr(&a));
                        }
                    }
                    "spcPct" if in_lst && in_ln_spc_lvl && cur_lvl.is_some() => {
                        if let (Some(idx), Some(v)) = (cur_lvl, get_attr(&e, "val")) {
                            if let Ok(x) = v.parse::<f32>() {
                                levels[idx].line_spacing = Some(x / 100000.0);
                            }
                        }
                    }
                    "lnSpc" if in_lst => in_ln_spc_lvl = true,
                    "latin" if in_lst && cur_lvl.is_some() => {
                        if let Some(idx) = cur_lvl {
                            if let Some(t) = get_attr(&e, "typeface") {
                                if !t.is_empty() && levels[idx].font_family.is_none() {
                                    levels[idx].font_family = Some(t);
                                }
                            }
                        }
                    }
                    "defRPr" if in_lst && cur_lvl.is_some() => {
                        in_def_rpr = true;
                        if let (Some(idx), Some(sz)) = (cur_lvl, get_attr(&e, "sz")) {
                            if let Ok(v) = sz.parse::<f32>() {
                                levels[idx].font_size = Some(v / 100.0);
                            }
                        }
                        if let (Some(idx), Some(i)) = (cur_lvl, get_attr(&e, "i")) {
                            if s_lvlitalic {
                                levels[idx].italic = i == "1" || i == "true";
                            }
                        }
                        // A level asks for WEIGHT the same way it asks for
                        // slant. d11's master title placeholder says b="1" and
                        // every title in the deck is bold; 461 placeholder
                        // levels over 27 dev decks declare one.
                        if let (Some(idx), Some(b)) = (cur_lvl, get_attr(&e, "b")) {
                            if s_lvlbold {
                                levels[idx].bold = Some(b == "1" || b == "true");
                            }
                        }
                    }
                    "highlight" if in_def_rpr => in_lvl_highlight = true,
                    "srgbClr" | "schemeClr" if in_def_rpr && in_lvl_highlight => {
                        if let (Some(idx), Some(val)) = (cur_lvl, get_attr(&e, "val")) {
                            if levels[idx].highlight.is_none() && s_highlight_lvl {
                                levels[idx].highlight = Some(if name == "srgbClr" {
                                    val
                                } else {
                                    theme_colors
                                        .get(&val)
                                        .cloned()
                                        .unwrap_or_else(|| scheme_color_to_hex(&val))
                                });
                            }
                        }
                    }
                    "srgbClr" | "schemeClr" if in_def_rpr => {
                        if let (Some(idx), Some(val)) = (cur_lvl, get_attr(&e, "val")) {
                            if levels[idx].color.is_none() {
                                levels[idx].color = Some(if name == "srgbClr" {
                                    val
                                } else {
                                    theme_colors
                                        .get(&val)
                                        .cloned()
                                        .unwrap_or_else(|| scheme_color_to_hex(&val))
                                });
                            }
                        }
                    }
                    _ => {}
                }
            }
            Ok(Event::End(e)) => match local_name(e.name().as_ref()).as_str() {
                "defRPr" => {
                    in_def_rpr = false;
                    in_lvl_highlight = false;
                }
                "highlight" if in_lvl_highlight => in_lvl_highlight = false,
                "lnSpc" => in_ln_spc_lvl = false,
                "lstStyle" => in_lst = false,
                "sp" => {
                    if in_sp && !levels.is_empty() {
                        out.insert((ph_type.clone(), ph_idx.clone()), std::mem::take(&mut levels));
                    }
                    in_sp = false;
                    cur_lvl = None;
                }
                _ => {}
            },
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
        buf.clear();
    }
    out
}

/// `p:presentation/@firstSlideNum` -- the number PowerPoint prints for the
/// FIRST slide of the deck. An absent attribute means 1.
///
/// Derived, not assumed (`gen_pptx_slidenum.py` + `read_pptx_slidenum.py`,
/// three 6-slide arms exported by PowerPoint itself): with the attribute
/// absent the deck prints 1..6, with `firstSlideNum="5"` it prints 5..10 and
/// with `"100"` it prints 100..105 -- so the printed number is the slide's
/// 1-based position plus this, less one. The same probe pins that the value
/// CACHED inside the field is ignored: a field holding `<a:t>777</a:t>` on
/// slide 3 printed 3 / 7 / 102 in the three arms.
fn parse_first_slide_num(pres_xml: &str) -> u32 {
    let mut reader = Reader::from_str(pres_xml);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                if local_name(e.name().as_ref()) == "presentation" {
                    return get_attr(&e, "firstSlideNum")
                        .and_then(|v| v.parse::<u32>().ok())
                        .unwrap_or(1);
                }
            }
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
        buf.clear();
    }
    1
}

/// Read `p:embeddedFontLst` and pull in every `.fntdata` part it names.
///
/// The list sits in `ppt/presentation.xml`; each `p:embeddedFont` carries one
/// `p:font/@typeface` plus up to four style children (`p:regular`, `p:bold`,
/// `p:italic`, `p:boldItalic`), each an `r:id` pointing at the part. A missing
/// or unreadable part is skipped rather than failing the parse -- a deck whose
/// fonts we cannot load still renders with substitutes.
fn parse_embedded_fonts(
    pres_xml: &str,
    rid_to_path: &HashMap<String, String>,
    archive: &mut OoxmlArchive,
) -> Vec<EmbeddedFont> {
    let mut out = Vec::new();
    let mut reader = Reader::from_str(pres_xml);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();
    let mut in_list = false;
    let mut typeface: Option<String> = None;
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                let name = local_name(e.name().as_ref());
                match name.as_str() {
                    "embeddedFontLst" => in_list = true,
                    "embeddedFont" if in_list => typeface = None,
                    "font" if in_list => typeface = get_attr(&e, "typeface"),
                    "regular" | "bold" | "italic" | "boldItalic" if in_list => {
                        let Some(face) = typeface.clone() else { continue };
                        let Some(r_id) = e
                            .attributes()
                            .flatten()
                            .find(|a| {
                                let k = std::str::from_utf8(a.key.as_ref()).unwrap_or("");
                                k == "r:id" || k.ends_with(":id")
                            })
                            .map(|a| String::from_utf8_lossy(&a.value).to_string())
                        else {
                            continue;
                        };
                        let Some(target) = rid_to_path.get(&r_id) else { continue };
                        let path = if let Some(stripped) = target.strip_prefix('/') {
                            stripped.to_string()
                        } else {
                            format!("ppt/{}", target)
                        };
                        let Ok(data) = archive.read_binary_part(&path) else { continue };
                        if data.is_empty() {
                            continue;
                        }
                        out.push(EmbeddedFont {
                            typeface: face,
                            bold: matches!(name.as_str(), "bold" | "boldItalic"),
                            italic: matches!(name.as_str(), "italic" | "boldItalic"),
                            data,
                        });
                    }
                    _ => {}
                }
            }
            Ok(Event::End(e)) => {
                if local_name(e.name().as_ref()) == "embeddedFontLst" {
                    break;
                }
            }
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
        buf.clear();
    }
    out
}

/// Where a slide sits in its deck.
///
/// The two travel together because a `slidenum` field prints their sum less
/// one, so splitting them into separate parameters only ever invites them to
/// be passed apart.
struct SlidePos {
    /// 1-based position in `p:sldIdLst` order -- what `Slide.index` carries.
    index: usize,
    /// `p:presentation/@firstSlideNum`; 1 when the deck does not state it.
    first_slide_num: u32,
}

/// Parse a single slide XML into shapes.
fn parse_slide(
    xml: &str,
    pos: SlidePos,
    archive: &mut OoxmlArchive,
    slide_rels_path: &str,
    master_ph_geoms: &HashMap<(Option<String>, Option<String>), (f32, f32, f32, f32)>,
    master_ph_anchors: &HashMap<(Option<String>, Option<String>), String>,
    theme_colors: &HashMap<String, String>,
) -> Result<Slide, PptxError> {
    let SlidePos { index: slide_index, first_slide_num } = pos;

    // Parse slide relationships for image resolution
    let rels = if let Ok(Some(rels_xml)) = archive.try_read_part(slide_rels_path) {
        parse_relationships(&rels_xml).unwrap_or_default()
    } else {
        Default::default()
    };

    // A schemeClr on a SLIDE resolves against the clrScheme of the theme that
    // the slide's master points at -- not the deck-level ppt/theme/theme1.xml.
    // A deck may ship several theme parts and the master picks one by
    // relationship: 17 of the 40 dev decks route their slides to theme2.xml, so
    // every schemeClr on those 365 slides was resolved against a theme the
    // slide does not use.  d10 slide 6 is the specimen: its <p:bg> asks for
    // dk1, theme2 defines dk1 = A5BEFD and PowerPoint paints it (164,189,252 at
    // all four corners of its own PDF export), while theme1 -- the part Oxi
    // read -- has dk1 = 000000, so the whole page came out black.
    //   Routing is per SLIDE, not per deck: d05 ships two masters and its slides
    // all reach slideMaster2 -> theme1.xml, so it is not affected at all.
    //   The colour MAP (p:clrMap / p:clrMapOvr) is a separate mechanism and is
    // not involved here: all 886 dev slides carry <a:masterClrMapping/>, and of
    // their 25160 schemeClr references not one names a mapped slot
    // (bg1/tx1/bg2/tx2) -- the histogram is dk1/lt1/lt2/dk2/accent1..6/hlink.
    //   Fonts stay on the deck-level theme: theme1 and the slide's own theme
    // agree on major/minor latin+ea in all 17 rerouted decks, so there is no
    // measurement to justify moving them.
    let slide_theme_colors = if std::env::var("OXI_SLIDETHEME_DISABLE").is_err() {
        resolve_slide_theme_colors(&rels, slide_rels_path, archive)
    } else {
        None
    };
    let theme_colors: &HashMap<String, String> =
        slide_theme_colors.as_ref().unwrap_or(theme_colors);

    // Spec #3: build a (ph_type, ph_idx) -> (x,y,w,h) geometry map from the
    // referenced slideLayout's placeholders. A slide placeholder with NO explicit
    // xfrm in its spPr inherits the layout placeholder's geometry. Spec #8
    // follow-up: when the layout placeholder also lacks an explicit xfrm, fall
    // back to the slideMaster's placeholder geometry (the master ph carries the
    // authoritative xfrm for layout-less placeholder slots).
    // The MASTER's placeholder lstStyles, merged UNDER the layout's.
    let master_ph_styles: HashMap<(Option<String>, Option<String>), Vec<MasterStyleLevel>> =
        if std::env::var("OXI_PHLEVEL_DISABLE").is_err() {
            match archive.try_read_part("ppt/slideMasters/slideMaster1.xml") {
                Ok(Some(xml)) => parse_layout_ph_lststyles(&xml, theme_colors),
                _ => HashMap::new(),
            }
        } else {
            HashMap::new()
        };

    let (layout_ph_geoms, layout_ph_anchors, layout_ph_styles): (
        HashMap<(Option<String>, Option<String>), (f32, f32, f32, f32)>,
        HashMap<(Option<String>, Option<String>), String>,
        HashMap<(Option<String>, Option<String>), Vec<MasterStyleLevel>>,
    ) = {
        let mut geoms = HashMap::new();
        let mut anchors = HashMap::new();
        let mut styles = HashMap::new();
        for rel in rels.values() {
            if rel.rel_type.ends_with("/slideLayout") {
                let layout_path =
                    resolve_slide_relative_path(slide_rels_path, &rel.target);
                if let Ok(Some(layout_xml)) = archive.try_read_part(&layout_path) {
                    let (g, a) = parse_layout_ph_info(&layout_xml).unwrap_or_default();
                    geoms = g;
                    anchors = a;
                    if std::env::var("OXI_PHLEVEL_DISABLE").is_err() {
                        styles = parse_layout_ph_lststyles(&layout_xml, theme_colors);
                    }
                }
                break;
            }
        }
        (geoms, anchors, styles)
    };

    // A slide with no <p:bg> of its own inherits the layout's background, and
    // the layout the master's -- PowerPoint paints that inherited fill edge to
    // edge.  Measured on the dev corpus: 24.5% of slides take their background
    // from the master and 6.2% from the layout, so without this 272 of 886
    // slides (30.7%) rendered on plain white.  d06 slide 34 is the specimen:
    // slide and layout both have no <p:bg>, slideMaster1 carries
    // <a:srgbClr val="77588B">, and PowerPoint's own PDF is that colour to the
    // page corners.
    let bgimg_on = std::env::var("OXI_BGIMG_DISABLE").is_err();
    // The shapes a layout/master draws for itself are inherited by the same
    // walk, but they are a separate concern from the background, so they have
    // their own switch. OXI_BGINHERIT_DISABLE turns the whole walk off, which
    // keeps it usable as "inherit nothing from layout/master".
    let lmshapes_on = std::env::var("OXI_LMSHAPES_DISABLE").is_err();
    let mut inherited_shapes: Vec<Shape> = Vec::new();
    let (inherited_bg, inherited_grad, inherited_img): (
        Option<String>,
        Option<SlideGradient>,
        Option<SlideBackgroundImage>,
    ) = if std::env::var("OXI_BGINHERIT_DISABLE").is_ok() {
        (None, None, None)
    } else {
        let mut found: Option<String> = None;
        let mut found_grad: Option<SlideGradient> = None;
        let mut found_img: Option<SlideBackgroundImage> = None;
        // `p:sld/@showMasterSp="0"` switches the master's own shapes off.
        let slide_hides_master = show_master_sp_off(xml);
        for rel in rels.values() {
            if !rel.rel_type.ends_with("/slideLayout") {
                continue;
            }
            let layout_path = resolve_slide_relative_path(slide_rels_path, &rel.target);
            if let Ok(Some(layout_xml)) = archive.try_read_part(&layout_path) {
                // layout -> master, and that master's THEME.  A schemeClr in a
                // layout/master resolves against the theme of ITS master, not
                // the deck-level theme1.xml: every dev deck ships one theme
                // part per master.  Measured -- d19 slideLayout5 asks for dk1,
                // which theme2.xml defines as 21355A but theme1.xml does not
                // define at all, so the deck-level map fell through to the
                // hardcoded black.  (d08 slides 6/10/13/26/27 are the same bug
                // in miniature: lt1 is EDECED in theme2, FFFFFF in theme1.)
                let cut = layout_path.rfind('/').map(|i| i + 1).unwrap_or(0);
                let layout_rels = format!(
                    "{}_rels/{}.rels",
                    &layout_path[..cut],
                    &layout_path[cut..]
                );
                let mut master_path: Option<String> = None;
                let mut master_colors: Option<HashMap<String, String>> = None;
                if let Ok(Some(lr_xml)) = archive.try_read_part(&layout_rels) {
                    if let Ok(lr) = parse_relationships(&lr_xml) {
                        for r in lr.values() {
                            if !r.rel_type.ends_with("/slideMaster") {
                                continue;
                            }
                            let mp = resolve_slide_relative_path(&layout_rels, &r.target);
                            let mcut = mp.rfind('/').map(|i| i + 1).unwrap_or(0);
                            let master_rels =
                                format!("{}_rels/{}.rels", &mp[..mcut], &mp[mcut..]);
                            if let Ok(Some(mr_xml)) = archive.try_read_part(&master_rels) {
                                if let Ok(mr) = parse_relationships(&mr_xml) {
                                    for t in mr.values() {
                                        if !t.rel_type.ends_with("/theme") {
                                            continue;
                                        }
                                        let tp = resolve_slide_relative_path(
                                            &master_rels,
                                            &t.target,
                                        );
                                        if let Ok(Some(tx)) = archive.try_read_part(&tp) {
                                            if let Ok((_, _, c)) = parse_theme(&tx) {
                                                master_colors = Some(c);
                                            }
                                        }
                                        break;
                                    }
                                }
                            }
                            master_path = Some(mp);
                            break;
                        }
                    }
                }
                let colors = master_colors.as_ref().unwrap_or(theme_colors);

                // The shapes a layout/master paints on its own account. Unlike
                // the background these do NOT stop at the layout: PowerPoint
                // draws the master's, then the layout's, then the slide's, so
                // both parts contribute and the order is master-first.
                if lmshapes_on {
                    if !slide_hides_master && !show_master_sp_off(&layout_xml) {
                        if let Some(mp) = master_path.as_ref() {
                            if let Ok(Some(mx)) = archive.try_read_part(mp) {
                                let mcut = mp.rfind('/').map(|i| i + 1).unwrap_or(0);
                                let master_rels =
                                    format!("{}_rels/{}.rels", &mp[..mcut], &mp[mcut..]);
                                inherited_shapes.extend(parse_inherited_shapes(
                                    &mx,
                                    &master_rels,
                                    archive,
                                    colors,
                                ));
                            }
                        }
                    }
                    inherited_shapes.extend(parse_inherited_shapes(
                        &layout_xml,
                        &layout_rels,
                        archive,
                        colors,
                    ));
                }

                found = parse_bg_solid_fill(&layout_xml, colors);
                found_grad = parse_bg_gradient(&layout_xml, colors);
                // Gated here as well as at the construction site: a layout that
                // declares ONLY a picture stops the walk below, so leaving this
                // enabled while the feature is off would turn such a slide
                // white instead of falling through to the master -- i.e. the
                // A/B "off" arm would not reproduce the pre-feature behaviour.
                found_img = if bgimg_on {
                    load_bg_image(&layout_xml, &layout_rels, archive)
                } else {
                    None
                };
                // A layout that declares a gradient or a picture HAS supplied
                // the background even though it yields no flat colour, so the
                // walk must stop there rather than fall through to the master.
                if found.is_none() && found_grad.is_none() && found_img.is_none() {
                    if let Some(mp) = master_path {
                        if let Ok(Some(mx)) = archive.try_read_part(&mp) {
                            let mcut = mp.rfind('/').map(|i| i + 1).unwrap_or(0);
                            let master_rels =
                                format!("{}_rels/{}.rels", &mp[..mcut], &mp[mcut..]);
                            found = parse_bg_solid_fill(&mx, colors);
                            found_grad = parse_bg_gradient(&mx, colors);
                            found_img = if bgimg_on {
                                load_bg_image(&mx, &master_rels, archive)
                            } else {
                                None
                            };
                        }
                    }
                }
            }
            break;
        }
        (found, found_grad, found_img)
    };

    let mut reader = Reader::from_str(xml);
    let mut shapes = Vec::new();
    // The editor's own numbering: `<p:sp>` and `<p:pic>` in document order.
    // Kept here so a shape carries the index an edit must be addressed by.
    let mut sp_count: usize = 0;
    let mut _depth = 0u32;
    let mut in_sp_tree = false;

    // Slide background state
    let mut in_bg = false;
    let mut in_bg_pr = false;
    let mut slide_background_color: Option<String> = None;
    // A slide that declares ANY <p:bg> owns its background outright -- it must
    // not fall back to the layout/master even when we cannot resolve the fill.
    // d19 slide 10 is the specimen: its own background is a blipFill (a paper
    // texture PowerPoint paints at #F2F1ED) which yields no single colour, and
    // inheriting instead put slideLayout5's dk1 (#21355A, near-black navy) over
    // the whole page.  47 dev slides own a grad/blip background this way.
    let mut slide_has_own_bg = false;

    // Shape state
    let mut in_shape = false;
    let mut shape_x: f32 = 0.0;
    let mut shape_y: f32 = 0.0;
    let mut shape_w: f32 = 0.0;
    let mut shape_h: f32 = 0.0;
    let mut shape_rotation: f32 = 0.0;
    let mut shape_flip_h = false;
    let mut shape_flip_v = false;
    let mut shape_prst: Option<String> = None;
    let mut in_prst_geom = false;
    let mut shape_adjustments: HashMap<String, f32> = HashMap::new();
    // a:custGeom — explicit outline paths. `cg_pending` is the command being
    // filled with its <a:pt> children (a cubicBezTo takes three of them, so the
    // points cannot be turned into a command until its end tag).
    // The highlight's opt-out has to reach the PARSER too: routing the colour
    // out of `run_color` is a content change, so a flag that only silences the
    // renderer leaves the "off" arm drawing the text in the highlight's colour
    // rather than reproducing the pre-change build. 8 decks / 16 pages proved it.
    let s_highlight = std::env::var("OXI_HIGHLIGHT_DISABLE").is_err();
    // `a:rPr/@spc` -- letter spacing. 60 runs over 3 blind decks ask for it.
    let s_letterspc = std::env::var("OXI_LETTERSPC_DISABLE").is_err();
    let s_custgeom = std::env::var("OXI_CUSTGEOM_DISABLE").is_err();
    let mut in_cust_geom = false;
    let mut cg_paths: Vec<GeomPath> = Vec::new();
    let mut cg_unsupported = false;
    let mut cg_cur: Option<GeomPath> = None;
    let mut cg_pending: Option<(&'static str, Vec<(f32, f32)>)> = None;
    let mut shape_paragraphs: Vec<SlideParagraph> = Vec::new();
    let mut shape_is_image = false;
    let mut shape_image_r_id: Option<String> = None;
    // Image fill geometry (a:blipFill): source crop (a:srcRect l/t/r/b) and
    // destination insets (a:stretch/a:fillRect l/t/r/b), normalized 0..1.
    // Word render-truth (01__Biology deck, 2026-08): a full-bleed background
    // PNG crops the SOURCE via srcRect; a photo expands the DESTINATION via a
    // negative fillRect so the stretched image keeps its native aspect.
    let mut shape_src_rect: Option<(f32, f32, f32, f32)> = None;
    let mut shape_rot_with_shape = true;
    let mut shape_image_alpha: Option<f32> = None;
    // a:gradFill on the shape itself (same ramp model as the background).
    let s_shapegrad = std::env::var("OXI_SHAPEGRAD_DISABLE").is_err();
    // `<a:lin>` / `<a:fillToRect>` inside a shape gradient are read from the
    // Empty branch (they are always self-closing) unless this is set, which
    // restores every shape ramp running at angle 0.
    let s_gradlin = std::env::var("OXI_GRADLIN_DISABLE").is_err();
    // A shape gradient's ramp centre comes from a:tileRect, not a:fillToRect.
    let s_gradtile = std::env::var("OXI_GRADTILE_DISABLE").is_err();
    // A gradient stop's colour is read from the Empty branch too unless this
    // is set, which restores dropping every stop written without an alpha.
    let s_gradstop = std::env::var("OXI_GRADSTOP_DISABLE").is_err();
    let mut sg_rot_with_shape = true;
    let mut sg_in = false;
    let mut sg_in_gs = false;
    let mut sg_in_path = false;
    let mut sg_pos: f32 = 0.0;
    let mut sg_color: Option<String> = None;
    let mut sg_alpha: f32 = 1.0;
    let mut sg_stops: Vec<SlideGradientStop> = Vec::new();
    let mut sg_angle: Option<f32> = None;
    let mut sg_scaled = false;
    let mut sg_focus: Option<(f32, f32)> = None;
    let mut shape_fill_rect: Option<(f32, f32, f32, f32)> = None;
    let mut shape_fill_color: Option<String> = None;
    let mut shape_fill_alpha: Option<f32> = None;
    let mut shape_border_color: Option<String> = None;
    let mut shape_border_alpha: Option<f32> = None;
    let mut shape_border_width: Option<f32> = None;
    let mut shape_border_dash: Option<String> = None;
    let mut shape_head_end: Option<LineEnd> = None;
    let mut shape_tail_end: Option<LineEnd> = None;
    let mut shape_line_cap: Option<String> = None;
    let mut shape_text_warp: Option<String> = None;
    // Placeholder identity (p:ph type/idx from nvPr) and whether spPr had an
    // explicit xfrm. Spec #3: a placeholder without an explicit xfrm inherits
    // its geometry from the referenced slideLayout's matching placeholder.
    let mut shape_ph_type: Option<String> = None;
    let mut shape_ph_idx: Option<String> = None;
    let mut shape_has_xfrm = false;
    // Group (p:grpSp) child-space transform stack.  Each entry is the
    // cumulative (offset_x, offset_y, scale_x, scale_y) mapping a child's
    // coordinates to slide points.  Empty = top level (identity).
    let s_grp = std::env::var("OXI_GRPXFRM_DISABLE").is_err();
    let s_grprot = std::env::var("OXI_GRPROT_DISABLE").is_err();
    // The mirror ships after the turn, so it needs its own opt-out: disabling
    // both together would compare against the pre-rotation build, not the
    // shipped one.
    let s_grpflip = std::env::var("OXI_GRPFLIP_DISABLE").is_err();
    // A NESTED group's own box turns and mirrors with its parent, as a
    // leaf shape already did. Its own flag, so the arm-A proof isolates it
    // from the leaf-level group flip that shipped before it.
    let s_grpnest = std::env::var("OXI_GRPNEST_DISABLE").is_err();
    // (origin x, origin y, scale x, scale y, rotation deg, centre x, centre y,
    //  flipH, flipV) -- the group's own turn and mirror, about that centre
    let mut grp_stack: Vec<(f32, f32, f32, f32, f32, f32, f32, bool, bool)> = Vec::new();
    let mut in_grp_sp_pr = false;
    let mut in_grp_xfrm = false;
    let mut g_off = (0.0f32, 0.0f32);
    let mut g_ext = (0.0f32, 0.0f32);
    let mut g_ch_off = (0.0f32, 0.0f32);
    let mut g_ch_ext = (0.0f32, 0.0f32);
    let mut g_rot: f32 = 0.0;
    let mut g_flip = (false, false);
    // Spec #6: vertical text-anchor from the shape's own a:bodyPr/@anchor
    // (resolved through the placeholder chain at shape end).
    let mut shape_anchor: Option<String> = None;
    let mut shape_wrap = true;
    let mut shape_spc_first_last = false;

    // Shape property context tracking
    let mut in_sp_pr = false; // inside <p:spPr> or <xdr:spPr>
    let mut in_ln = false;    // inside <a:ln> (line/border properties)
    // Inside <a:effectLst>. Its colours belong to the shadow / glow, not to the
    // shape, but it is a child of p:spPr like a:solidFill is -- so without this
    // the LAST colour in spPr wins and a drop shadow repaints the shape.
    // d06 slide 14 is a white world map with a black `a:outerShdw`, and Oxi drew
    // the continents BLACK; 110 shapes across seven decks are shaped like that.
    let mut in_effect_lst = false;
    let s_effectclr = std::env::var("OXI_EFFECTCLR_DISABLE").is_err();
    // Inside <a:outerShdw>, whose colour and alpha are the SHADOW's.
    let mut in_outer_shdw = false;
    // (blur pt, dist pt, dir deg) while the element is open.
    let mut shape_shadow_draft: Option<(f32, f32, f32)> = None;
    let mut shape_shadow_color: Option<String> = None;
    let mut shape_shadow_alpha: f32 = 1.0;
    let mut shape_shadow: Option<crate::ir::ShapeShadow> = None;
    // Inside <a:solidFill>. `a:alpha` also appears under a gradient stop and a
    // line colour, and only a solid shape fill is composited (S-FILLALPHA).
    let mut in_solid_fill = false;

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
    // `a:lnSpc/a:spcPts` -- an exact line height, overriding the multiple.
    let s_lnspcpts = std::env::var("OXI_LNSPCPTS_DISABLE").is_err();
    let mut para_line_spacing_pts: Option<f32> = None;
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
    let mut para_end_size: Option<f32> = None;

    // Run state
    let mut in_run = false;
    // `<a:fld>` is a run that carries a FIELD rather than literal text: the
    // `<a:t>` inside it is only PowerPoint's cache of what the field last
    // printed, and PowerPoint recomputes it on every render. It is parsed
    // exactly like `<a:r>` -- same rPr, same a:t -- with the text substituted
    // when the element closes. `None` here means the current run is a plain
    // `<a:r>`. See `parse_first_slide_num` for the derivation.
    let mut fld_type: Option<String> = None;
    let s_slidenum = std::env::var("OXI_SLIDENUM_DISABLE").is_err();
    let mut run_text = String::new();
    let mut run_bold: Option<bool> = None;
    let mut run_italic = false;
    let mut run_underline = false;
    let mut run_font_size: Option<f32> = None;
    let mut run_spacing: Option<f32> = None;
    let mut run_color: Option<String> = None;
    let mut run_color_alpha: Option<f32> = None;
    // `a:rPr/a:highlight` -- the run's text highlight. It holds a colour
    // element of exactly the shape `a:solidFill` does, so without this flag
    // the colour dispatch below reads it as the run's own text colour: d11
    // slide 38's "and many more..." is white on dk1 and Oxi drew it dk1 on
    // nothing.
    let s_softbreak = std::env::var("OXI_SOFTBREAK_DISABLE").is_err();
    let mut in_highlight = false;
    let mut run_highlight: Option<String> = None;
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
    let mut tc_grid_span: u32 = 1;
    let mut tc_h_merge = false;
    // a:tcPr — the cell's own fill, per-side borders, margins and anchor. The
    // corpus states all of them explicitly (an INVISIBLE edge is written as a
    // solidFill with <a:alpha val="0"/>, not as an absent element), so none of
    // this needs the table style resolved.
    let mut in_tc_pr = false;
    let mut tc_ln_side: Option<usize> = None;
    let mut tc_borders: [Option<CellBorder>; 4] = Default::default();
    let mut tc_fill: Option<String> = None;
    let mut tc_fill_alpha: Option<f32> = None;
    let mut tc_mar = (default_l_ins(), default_r_ins(), default_t_ins(), default_b_ins());
    let mut tc_anchor: Option<String> = None;
    let mut tc_ln_width: f32 = 0.0;
    let mut tc_ln_color: Option<String> = None;
    let mut tc_ln_alpha: f32 = 1.0;

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
                        slide_has_own_bg = true;
                    }
                    "bgPr" if in_bg => {
                        in_bg_pr = true;
                    }
                    "spTree" => {
                        in_sp_tree = true;
                    }
                    "grpSp" if in_sp_tree => {
                        // Inherit the parent transform until this group's
                        // own a:xfrm is read (a group may omit it).
                        let base =
                            grp_stack
                                .last()
                                .copied()
                                .unwrap_or((0.0, 0.0, 1.0, 1.0, 0.0, 0.0, 0.0, false, false));
                        grp_stack.push(base);
                    }
                    "grpSpPr" if !grp_stack.is_empty() => {
                        in_grp_sp_pr = true;
                    }
                    "xfrm" if in_grp_sp_pr => {
                        in_grp_xfrm = true;
                        g_off = (0.0, 0.0);
                        g_ext = (0.0, 0.0);
                        g_ch_off = (0.0, 0.0);
                        g_ch_ext = (0.0, 0.0);
                        // `p:grpSpPr/a:xfrm/@rot` turns the whole group about
                        // its own box centre. d19 slide 36 is ten pencils, each
                        // a 25x381pt picture inside its own group at rot 90 --
                        // ignoring it drew ten vertical pencils stacked at one
                        // x instead of ten horizontal ones down the slide.
                        g_rot = get_attr(&e, "rot")
                            .and_then(|v| v.parse::<f32>().ok())
                            .map(|v| v / 60000.0)
                            .unwrap_or(0.0);
                        g_flip = (
                            get_attr(&e, "flipH").as_deref() == Some("1"),
                            get_attr(&e, "flipV").as_deref() == Some("1"),
                        );
                    }
                    "sp" | "pic" | "cxnSp" if in_sp_tree => {
                        in_shape = true;
                        shape_x = 0.0;
                        shape_y = 0.0;
                        shape_w = 0.0;
                        shape_h = 0.0;
                        shape_rotation = 0.0;
                        shape_flip_h = false;
                        shape_flip_v = false;
                        shape_prst = None;
                        in_prst_geom = false;
                        shape_adjustments.clear();
                        cg_paths.clear();
                        cg_unsupported = false;
                        cg_cur = None;
                        cg_pending = None;
                        in_cust_geom = false;
                        shape_paragraphs.clear();
                        shape_is_image = name == "pic";
                        shape_image_r_id = None;
                        shape_src_rect = None;
                        shape_rot_with_shape = true;
                        shape_image_alpha = None;
                        sg_in = false;
                        sg_in_gs = false;
                        sg_in_path = false;
                        sg_stops.clear();
                        sg_angle = None;
                        sg_scaled = false;
                        sg_focus = None;
                        shape_fill_rect = None;
                        shape_fill_color = None;
                        shape_fill_alpha = None;
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
                        shape_wrap = true;
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
                        shape_flip_h = false;
                        shape_flip_v = false;
                        shape_prst = None;
                        in_prst_geom = false;
                        shape_adjustments.clear();
                        cg_paths.clear();
                        cg_unsupported = false;
                        cg_cur = None;
                        cg_pending = None;
                        in_cust_geom = false;
                        shape_paragraphs.clear();
                        shape_is_image = false;
                        shape_image_r_id = None;
                        shape_fill_color = None;
                        shape_fill_alpha = None;
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
                        shape_wrap = true;
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
                        // `gridSpan` / `hMerge` live on a:tc itself. Both halves
                        // matter: the spanning cell needs the width of all the
                        // columns it covers, and each continuation cell still
                        // occupies a grid slot that must not be painted again.
                        tc_grid_span = get_attr(&e, "gridSpan")
                            .and_then(|v| v.parse::<u32>().ok())
                            .unwrap_or(1)
                            .max(1);
                        tc_h_merge = matches!(
                            get_attr(&e, "hMerge").as_deref(),
                            Some("1") | Some("true")
                        );
                    }
                    "tcPr" if in_tbl_cell => {
                        in_tc_pr = true;
                        tc_borders = Default::default();
                        tc_fill = None;
                        tc_fill_alpha = None;
                        tc_ln_side = None;
                        tc_mar = (
                            get_attr(&e, "marL").and_then(|v| v.parse::<f32>().ok()).map(emu_to_pt).unwrap_or_else(default_l_ins),
                            get_attr(&e, "marR").and_then(|v| v.parse::<f32>().ok()).map(emu_to_pt).unwrap_or_else(default_r_ins),
                            get_attr(&e, "marT").and_then(|v| v.parse::<f32>().ok()).map(emu_to_pt).unwrap_or_else(default_t_ins),
                            get_attr(&e, "marB").and_then(|v| v.parse::<f32>().ok()).map(emu_to_pt).unwrap_or_else(default_b_ins),
                        );
                        tc_anchor = get_attr(&e, "anchor");
                    }
                    "lnL" | "lnR" | "lnT" | "lnB" if in_tc_pr => {
                        tc_ln_side = Some(match name.as_str() {
                            "lnL" => 0,
                            "lnR" => 1,
                            "lnT" => 2,
                            _ => 3,
                        });
                        tc_ln_width = get_attr(&e, "w")
                            .and_then(|v| v.parse::<f32>().ok())
                            .map(|v| v / 12700.0)
                            .unwrap_or(0.75);
                        tc_ln_color = None;
                        tc_ln_alpha = 1.0;
                    }
                    "srgbClr" | "schemeClr" if in_tc_pr => {
                        if let Some(val) = get_attr(&e, "val") {
                            let hex = if name == "srgbClr" {
                                val
                            } else {
                                theme_colors
                                    .get(&val)
                                    .cloned()
                                    .unwrap_or_else(|| scheme_color_to_hex(&val))
                            };
                            if tc_ln_side.is_some() {
                                tc_ln_color = Some(hex);
                            } else {
                                tc_fill = Some(hex);
                            }
                        }
                    }
                    // A run's own colour alpha (d35's 26.9% white numerals).
                    "alpha" if in_outer_shdw => {
                        if let Some(v) = get_attr(&e, "val").and_then(|v| v.parse::<f32>().ok()) {
                            shape_shadow_alpha = (v / 100000.0).clamp(0.0, 1.0);
                        }
                    }
                    "alpha" if in_ln && in_sp_pr => {
                        // The BORDER's translucency (a:ln/a:solidFill/a:alpha).
                        // d49's site pill is a 3pt white ring at 35.3% over
                        // black; opaque it reads as a chalk outline.
                        if let Some(v) = get_attr(&e, "val").and_then(|v| v.parse::<f32>().ok()) {
                            shape_border_alpha = Some((v / 100000.0).clamp(0.0, 1.0));
                        }
                    }
                    "alpha" if in_run && !in_highlight => {
                        if let Some(v) = get_attr(&e, "val") {
                            if let Ok(pc) = v.parse::<f32>() {
                                run_color_alpha = Some((pc / 100000.0).clamp(0.0, 1.0));
                            }
                        }
                    }
                    "alpha" if in_tc_pr => {
                        if let Some(v) = get_attr(&e, "val") {
                            if let Ok(p) = v.parse::<f32>() {
                                let a = (p / 100000.0).clamp(0.0, 1.0);
                                if tc_ln_side.is_some() {
                                    tc_ln_alpha = a;
                                } else {
                                    tc_fill_alpha = Some(a);
                                }
                            }
                        }
                    }
                    "xfrm" if in_shape => {
                        // a:xfrm/@rot is in 60000ths of a degree; 60000 = 1 degree.
                        shape_has_xfrm = true;
                        shape_flip_h = get_attr(&e, "flipH").as_deref() == Some("1");
                        shape_flip_v = get_attr(&e, "flipV").as_deref() == Some("1");
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
                        in_prst_geom = true;
                    }
                    "gd" if in_prst_geom => {
                        if let (Some(name), Some(fmla)) = (get_attr(&e, "name"), get_attr(&e, "fmla")) {
                            if let Some(value) = fmla
                                .strip_prefix("val ")
                                .and_then(|v| v.trim().parse::<f32>().ok())
                            {
                                shape_adjustments.insert(name, value);
                            }
                        }
                    }
                    // a:gradFill on the SHAPE (not the background, not inside
                    // a:ln). Same ramp model as the slide background, which is
                    // already derived; only the paint area differs.
                    "gradFill" if in_sp_pr && !in_ln && s_shapegrad => {
                        sg_in = true;
                        sg_stops.clear();
                        sg_angle = None;
                        sg_scaled = false;
                        sg_focus = None;
                        // Absent behaves as "1" -- measured, not assumed
                        // (`gradrot` probe block B, absent == rws "1" on 6 of 6
                        // arms while rws "0" pinned the ramp to the page).
                        sg_rot_with_shape = get_attr(&e, "rotWithShape").as_deref() != Some("0");
                    }
                    "gs" if sg_in => {
                        sg_in_gs = true;
                        sg_pos = get_attr(&e, "pos").and_then(|v| gradient_frac(&v)).unwrap_or(0.0);
                        sg_color = None;
                        sg_alpha = 1.0;
                    }
                    "alpha" if sg_in_gs => {
                        if let Some(v) = get_attr(&e, "val") {
                            if let Ok(a) = v.parse::<f32>() {
                                sg_alpha = (a / 100000.0).clamp(0.0, 1.0);
                            }
                        }
                    }
                    "srgbClr" if sg_in_gs && sg_color.is_none() => {
                        sg_color = get_attr(&e, "val");
                    }
                    "schemeClr" if sg_in_gs && sg_color.is_none() => {
                        if let Some(v) = get_attr(&e, "val") {
                            sg_color = Some(
                                theme_colors
                                    .get(&v)
                                    .cloned()
                                    .unwrap_or_else(|| scheme_color_to_hex(&v)),
                            );
                        }
                    }
                    "lin" if sg_in => {
                        sg_angle = get_attr(&e, "ang")
                            .and_then(|v| v.parse::<f32>().ok())
                            .map(|v| v / 60_000.0);
                        sg_scaled = get_attr(&e, "scaled").as_deref() == Some("1");
                    }
                    "path" if sg_in => {
                        if get_attr(&e, "path").as_deref() == Some("circle") {
                            sg_in_path = true;
                            sg_focus = Some((0.5, 0.5));
                        }
                    }
                    // S-GRADTILE (2026-08-27). For a SHAPE fill PowerPoint
                    // IGNORES `a:fillToRect` -- the gradpath probe renders the
                    // focus declared centre, top-left, bottom-right and ABSENT
                    // as four identical centred circles -- and takes the ramp's
                    // centre from `a:tileRect` instead. Reading fillToRect put
                    // the focus at (0.25,0.25) for d15-style `r=b=100%` (l and t
                    // default to 0.5 here), which is neither what PowerPoint
                    // does with the attribute nor what it does without it, and
                    // cost 0.2118 mean|delta| on the probe's arm 2.
                    "fillToRect" if sg_in_path && !s_gradtile => {
                        let l = get_attr(&e, "l").and_then(|v| gradient_frac(&v)).unwrap_or(0.5);
                        let t = get_attr(&e, "t").and_then(|v| gradient_frac(&v)).unwrap_or(0.5);
                        let r = get_attr(&e, "r").and_then(|v| gradient_frac(&v)).unwrap_or(0.5);
                        let b = get_attr(&e, "b").and_then(|v| gradient_frac(&v)).unwrap_or(0.5);
                        sg_focus = Some(((l + (1.0 - r)) / 2.0, (t + (1.0 - b)) / 2.0));
                    }
                    // The tile's own centre, in shape-relative units: its edges
                    // are INSETS from the shape box, so a negative one pushes
                    // outward. `l="-100%" t="-100%"` makes a tile twice the
                    // shape centred on the shape's TOP-LEFT corner, and that is
                    // where PowerPoint puts the ramp's centre (probe arm 4,
                    // focus (0.00,0.00)). Absent or empty -> (0.5,0.5), which is
                    // the centred case every other arm shows.
                    "tileRect" if sg_in && sg_focus.is_some() && s_gradtile => {
                        let l = get_attr(&e, "l").and_then(|v| gradient_frac(&v)).unwrap_or(0.0);
                        let t = get_attr(&e, "t").and_then(|v| gradient_frac(&v)).unwrap_or(0.0);
                        let r = get_attr(&e, "r").and_then(|v| gradient_frac(&v)).unwrap_or(0.0);
                        let b = get_attr(&e, "b").and_then(|v| gradient_frac(&v)).unwrap_or(0.0);
                        sg_focus = Some(((l + (1.0 - r)) / 2.0, (t + (1.0 - b)) / 2.0));
                    }
                    "custGeom" if in_shape && s_custgeom => {
                        in_cust_geom = true;
                        cg_paths.clear();
                        cg_unsupported = false;
                        cg_cur = None;
                        cg_pending = None;
                    }
                    "path" if in_cust_geom => {
                        cg_cur = Some(GeomPath {
                            w: get_attr(&e, "w")
                                .and_then(|v| v.parse::<f32>().ok())
                                .unwrap_or(0.0),
                            h: get_attr(&e, "h")
                                .and_then(|v| v.parse::<f32>().ok())
                                .unwrap_or(0.0),
                            fill_none: get_attr(&e, "fill").as_deref() == Some("none"),
                            commands: Vec::new(),
                        });
                    }
                    "moveTo" | "lnTo" | "cubicBezTo" if in_cust_geom => {
                        let kind = match name.as_str() {
                            "moveTo" => "moveTo",
                            "lnTo" => "lnTo",
                            _ => "cubicBezTo",
                        };
                        cg_pending = Some((kind, Vec::new()));
                    }
                    // Outside the measured vocabulary: refuse the whole geometry
                    // rather than draw an outline with a segment missing.
                    "arcTo" | "quadBezTo" if in_cust_geom => {
                        cg_unsupported = true;
                    }
                    "close" if in_cust_geom => {
                        if let Some(p) = cg_cur.as_mut() {
                            p.commands.push(GeomCmd::Close);
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
                    "effectLst" => {
                        in_effect_lst = true;
                    }
                    "outerShdw" if in_sp_pr && in_effect_lst => {
                        in_outer_shdw = true;
                        let emu = |name: &str| {
                            get_attr(&e, name)
                                .and_then(|v| v.parse::<f32>().ok())
                                .unwrap_or(0.0)
                        };
                        shape_shadow_draft =
                            Some((emu("blurRad") / 12700.0, emu("dist") / 12700.0, emu("dir") / 60000.0));
                        shape_shadow_color = None;
                        shape_shadow_alpha = 1.0;
                    }
                    "ln" if in_sp_pr => {
                        in_ln = true;
                        shape_border_dash = None;
                        shape_head_end = None;
                        shape_tail_end = None;
                        shape_line_cap = get_attr(&e, "cap");
                        // Width attribute in EMU; 12700 EMU = 1pt
                        if let Some(w) = get_attr(&e, "w") {
                            if let Ok(v) = w.parse::<f32>() {
                                shape_border_width = Some(v / 12700.0);
                            }
                        }
                    }
                    // `a:prstDash` is self-closing, so it arrives on the Empty
                    // arm as well -- both are routed here.
                    "prstDash" if in_ln => {
                        shape_border_dash = get_attr(&e, "val").filter(|v| v != "solid");
                    }
                    // Likewise `a:headEnd` / `a:tailEnd`: attribute-only, so
                    // the Empty arm is the one that fires in practice. The
                    // `in_ln` guard keeps table-cell borders (a:lnL/lnR/lnT/
                    // lnB, which state type="none" ends of their own) out.
                    "headEnd" | "tailEnd" if in_ln => {
                        let end = parse_line_end(&e);
                        if name == "headEnd" {
                            shape_head_end = end;
                        } else {
                            shape_tail_end = end;
                        }
                    }
                    "off" if in_grp_xfrm => {
                        if let Some(x) = get_attr(&e, "x") {
                            g_off.0 = x.parse::<f32>().map(emu_to_pt).unwrap_or(0.0);
                        }
                        if let Some(y) = get_attr(&e, "y") {
                            g_off.1 = y.parse::<f32>().map(emu_to_pt).unwrap_or(0.0);
                        }
                    }
                    "chOff" if in_grp_xfrm => {
                        if let Some(x) = get_attr(&e, "x") {
                            g_ch_off.0 = x.parse::<f32>().map(emu_to_pt).unwrap_or(0.0);
                        }
                        if let Some(y) = get_attr(&e, "y") {
                            g_ch_off.1 = y.parse::<f32>().map(emu_to_pt).unwrap_or(0.0);
                        }
                    }
                    "chExt" if in_grp_xfrm => {
                        if let Some(cx) = get_attr(&e, "cx") {
                            g_ch_ext.0 = cx.parse::<f32>().map(emu_to_pt).unwrap_or(0.0);
                        }
                        if let Some(cy) = get_attr(&e, "cy") {
                            g_ch_ext.1 = cy.parse::<f32>().map(emu_to_pt).unwrap_or(0.0);
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
                    "ext" if in_grp_xfrm => {
                        if let Some(cx) = get_attr(&e, "cx") {
                            g_ext.0 = cx.parse::<f32>().map(emu_to_pt).unwrap_or(0.0);
                        }
                        if let Some(cy) = get_attr(&e, "cy") {
                            g_ext.1 = cy.parse::<f32>().map(emu_to_pt).unwrap_or(0.0);
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
                        // @rotWithShape="0" pins the raster upright while the
                        // shape still turns (PowerPoint render-truth, E5).
                        if get_attr(&e, "rotWithShape").as_deref() == Some("0") {
                            shape_rot_with_shape = false;
                        }
                    }
                    // <a:alphaModFix amt=".."/> is self-closing, so it arrives
                    // in the Empty arm; both arms carry the handler because a
                    // producer may write either form.
                    "alphaModFix" if in_shape && shape_is_image => {
                        if let Some(v) = get_attr(&e, "amt") {
                            if let Ok(a) = v.parse::<f32>() {
                                shape_image_alpha = Some((a / 100000.0).clamp(0.0, 1.0));
                            }
                        }
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
                        para_line_spacing_pts = None;
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
                        para_end_size = None;
                    }
                    // `a:prstTxWarp` is normally self-closing, so both arms
                    // route here (the Start/Empty trap this file keeps).
                    "prstTxWarp" if in_body_pr => {
                        shape_text_warp = get_attr(&e, "prst");
                    }
                    "bodyPr" if in_shape => {
                        in_body_pr = true;
                        shape_text_warp = None;
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
                        // a:bodyPr/@wrap="none" — the text runs past the box
                        // instead of breaking. The `embedsplit` probes are
                        // COM-built and every arm's box is auto-sized to
                        // 14.5pt around 20pt text, so a renderer that wraps
                        // them anyway shreds "AAABBB" into six lines where
                        // PowerPoint draws one.
                        if let Some(w) = get_attr(&e, "wrap") {
                            shape_wrap = w != "none";
                        }
                        // `a:bodyPr/@spcFirstLastPara` -- keep the first
                        // paragraph's spcBef instead of dropping it. See
                        // `Shape::spc_first_last_para` for the probe that
                        // settled what it does.
                        if let Some(v) = get_attr(&e, "spcFirstLastPara") {
                            shape_spc_first_last = v == "1" || v == "true";
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
                    "r" | "fld" if in_paragraph => {
                        in_run = true;
                        fld_type = if name == "fld" {
                            get_attr(&e, "type").or_else(|| Some(String::new()))
                        } else {
                            None
                        };
                        run_text.clear();
                        run_bold = None;
                        run_italic = false;
                        run_underline = false;
                        run_font_size = None;
                        run_spacing = None;
                        run_color = None;
                        run_color_alpha = None;
                        run_highlight = None;
                        run_font_family = None;
                    }
                    "endParaRPr" if in_paragraph => {
                        para_end_size = get_attr(&e, "sz")
                            .and_then(|v| v.parse::<f32>().ok())
                            .map(|v| v / 100.0);
                    }
                    "rPr" if in_run => {
                        if let Some(b) = get_attr(&e, "b") {
                            run_bold = Some(b == "1" || b == "true");
                        }
                        if let Some(i) = get_attr(&e, "i") {
                            run_italic = i == "1" || i == "true";
                        }
                        if let Some(u) = get_attr(&e, "u") {
                            run_underline = u != "none";
                        }
                        if let Some(sz) = get_attr(&e, "sz") {
                            // Font size in hundredths of a point
                            if let Ok(v) = sz.parse::<f32>() {
                                run_font_size = Some(v / 100.0);
                            }
                        }
                        if s_letterspc {
                            if let Some(spc) = get_attr(&e, "spc") {
                                // Hundredths of a point, and may be negative.
                                if let Ok(v) = spc.parse::<f32>() {
                                    run_spacing = Some(v / 100.0);
                                }
                            }
                        }
                    }
                    // container — context determines where color goes
                    "solidFill" => in_solid_fill = true,
                    "highlight" if in_run => in_highlight = true,
                    "srgbClr" => {
                        if let Some(val) = get_attr(&e, "val") {
                            if in_outer_shdw {
                                shape_shadow_color = Some(val);
                            } else if in_bg_pr {
                                slide_background_color = Some(val);
                            } else if in_ln && in_sp_pr {
                                shape_border_color = Some(val);
                            } else if in_sp_pr && !in_ln && !(in_effect_lst && s_effectclr) {
                                shape_fill_color = Some(val);
                            } else if in_run && in_highlight && s_highlight {
                                run_highlight = Some(val);
                            } else if in_run {
                                run_color = Some(val);
                            }
                        }
                    }
                    "schemeClr" => {
                        if let Some(val) = get_attr(&e, "val") {
                            let hex = theme_colors.get(&val).cloned().unwrap_or_else(|| scheme_color_to_hex(&val));
                            if in_outer_shdw {
                                shape_shadow_color = Some(hex);
                            } else if in_bg_pr {
                                slide_background_color = Some(hex);
                            } else if in_ln && in_sp_pr {
                                shape_border_color = Some(hex);
                            } else if in_sp_pr && !in_ln && !(in_effect_lst && s_effectclr) {
                                shape_fill_color = Some(hex);
                            } else if in_run && in_highlight && s_highlight {
                                run_highlight = Some(hex);
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
                    // `<a:br/>` is a soft line break INSIDE a paragraph, not a
                    // new paragraph: it breaks the line without the paragraph's
                    // spcBef / spcAft. It is carried as a run holding a single
                    // newline so the wrap can honour it and the run's own
                    // properties still describe the line it ends. d19 slide 39's
                    // instructions are one paragraph with three of them, and
                    // ignoring them ran "quality." into "How?" with no space.
                    // 76 across 11 dev decks.
                    "br" if in_paragraph && s_softbreak => {
                        para_runs.push(SlideRun {
                            text: "
".to_string(),
                            font_size: run_font_size,
                            bold: None,
                            italic: false,
                            underline: false,
                            color: None,
                            color_alpha: None,
                            highlight: None,
                            font_family: run_font_family.clone(),
                            spacing: run_spacing,
                        });
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
                    // A SELF-CLOSING outerShdw is legal (nothing but
                    // attributes); the schema's default colour is black. The
                    // Start-arm handler never sees it, so it completes here in
                    // one step.
                    "outerShdw" if in_sp_pr && in_effect_lst => {
                        let emu = |name: &str| {
                            get_attr(&e, name)
                                .and_then(|v| v.parse::<f32>().ok())
                                .unwrap_or(0.0)
                        };
                        shape_shadow = Some(crate::ir::ShapeShadow {
                            blur_pt: emu("blurRad") / 12700.0,
                            dist_pt: emu("dist") / 12700.0,
                            dir_deg: emu("dir") / 60000.0,
                            color: "000000".to_string(),
                            alpha: 1.0,
                        });
                    }
                    // <a:alpha> only ever arrives self-closing (the schema gives
                    // it nothing but @val), which is why a gradient STOP's alpha
                    // has to be read here too. d15's illustrations are white at
                    // 20%/30% fading to 0 over a purple slide; read as opaque
                    // they paint a white slab.
                    "alpha" if sg_in_gs => {
                        if let Some(v) = get_attr(&e, "val") {
                            if let Ok(a) = v.parse::<f32>() {
                                sg_alpha = (a / 100000.0).clamp(0.0, 1.0);
                            }
                        }
                    }
                    "alpha" if in_solid_fill && in_sp_pr && !in_ln => {
                        if let Some(v) = get_attr(&e, "val") {
                            if let Ok(p) = v.parse::<f32>() {
                                shape_fill_alpha = Some((p / 100000.0).clamp(0.0, 1.0));
                            }
                        }
                    }
                    // S-GRADSTOP (2026-08-24): a stop whose colour carries no
                    // child -- `<a:gs pos="0"><a:srgbClr val="000000"/></a:gs>`
                    // -- reaches quick-xml on the Empty arm, and only the Start
                    // arm was reading it. The stop was dropped, and a ramp left
                    // with fewer than two stops is discarded whole, so the shape
                    // fell back to a flat fill or to nothing at all.
                    //
                    // **386 stops over 12 of the 40 dev decks are written this
                    // way, and 178 gradFills lose their second stop because of
                    // it** (d24 85, d15 32, d16 29, d09 23). It is invisible in
                    // the corpus decks that DO have alpha on every stop, which
                    // is why the ramp work up to here never tripped over it: an
                    // `<a:srgbClr>` with an `<a:alpha>` inside is a Start
                    // element. The gradrot probe -- authored with bare colours
                    // -- rendered as a blank page and gave it away.
                    //
                    // Same Start/Empty split as prstDash, prstTxWarp, run alpha,
                    // the line ends, the cell properties and a:lin before it.
                    "srgbClr" if sg_in_gs && sg_color.is_none() && s_gradstop => {
                        sg_color = get_attr(&e, "val");
                    }
                    "schemeClr" if sg_in_gs && sg_color.is_none() && s_gradstop => {
                        if let Some(v) = get_attr(&e, "val") {
                            sg_color = Some(
                                theme_colors
                                    .get(&v)
                                    .cloned()
                                    .unwrap_or_else(|| scheme_color_to_hex(&v)),
                            );
                        }
                    }
                    // A field with no cached `<a:t>` is self-closing, so it
                    // reaches quick-xml on the Empty arm and never on Start --
                    // the same split that has cost this renderer prstDash,
                    // prstTxWarp, run alpha and the line ends. It still prints:
                    // the number does not come from the file.
                    "fld" if in_paragraph && s_slidenum => {
                        if get_attr(&e, "type").as_deref() == Some("slidenum") {
                            para_runs.push(SlideRun {
                                text: (slide_index as u32 + first_slide_num - 1).to_string(),
                                font_size: run_font_size,
                                bold: None,
                                italic: false,
                                underline: false,
                                color: None,
                                color_alpha: None,
                                highlight: None,
                                font_family: run_font_family.clone(),
                                spacing: run_spacing,
                            });
                        }
                    }
                    // `<a:br/>` carries `<a:rPr>` only when something styles the
                    // break, and `rPr` is optional in the schema -- so the same
                    // element reaches Start in one file and Empty in another.
                    // Every deck in the corpus writes the Start form (149 of
                    // them, 0 self-closing), which is why this never showed
                    // there; a deck written by python-pptx is all Empty, and
                    // its paragraphs came back as ONE line with the breaks
                    // silently gone. Found by a probe that could not tell its
                    // seven arms apart; the pre-change binary run beside this
                    // one over dev + blind agrees on all 48365 paragraphs.
                    "br" if in_paragraph && s_softbreak => {
                        para_runs.push(SlideRun {
                            text: "\n".to_string(),
                            font_size: run_font_size,
                            bold: None,
                            italic: false,
                            underline: false,
                            color: None,
                            color_alpha: None,
                            highlight: None,
                            font_family: run_font_family.clone(),
                            spacing: run_spacing,
                        });
                    }
                    // `<a:endParaRPr sz="1400"/>` normally has no children, so it
                    // arrives here and never in the Start arm. The Start arm has
                    // its own copy for the form that wraps an a:solidFill.
                    "endParaRPr" if in_paragraph => {
                        para_end_size = get_attr(&e, "sz")
                            .and_then(|v| v.parse::<f32>().ok())
                            .map(|v| v / 100.0);
                    }
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
                    // ★`<a:lin>` and `<a:fillToRect>` are ALWAYS self-closing,
                    // so the Start-branch handlers for them never ran and every
                    // shape gradient in the corpus came out at angle 0 -- the
                    // ramp mirrored. d15 s17's three process bands are white
                    // washes at `ang="10800025"` (180 degrees), brightest at
                    // each shape's RIGHT edge in PowerPoint and at its LEFT in
                    // Oxi. 302 shape-level gradients over 4 dev decks. Same trap
                    // as the cell properties below, and the third time this file
                    // has paid for it.
                    "lin" if sg_in && s_gradlin => {
                        sg_angle = get_attr(&e, "ang")
                            .and_then(|v| v.parse::<f32>().ok())
                            .map(|v| v / 60_000.0);
                        sg_scaled = get_attr(&e, "scaled").as_deref() == Some("1");
                    }
                    "path" if sg_in && s_gradlin => {
                        if get_attr(&e, "path").as_deref() == Some("circle") {
                            sg_focus = Some((0.5, 0.5));
                        }
                    }
                    "fillToRect" if sg_in_path && s_gradlin && !s_gradtile => {
                        let l = get_attr(&e, "l").and_then(|v| gradient_frac(&v)).unwrap_or(0.5);
                        let t = get_attr(&e, "t").and_then(|v| gradient_frac(&v)).unwrap_or(0.5);
                        let r = get_attr(&e, "r").and_then(|v| gradient_frac(&v)).unwrap_or(0.5);
                        let b = get_attr(&e, "b").and_then(|v| gradient_frac(&v)).unwrap_or(0.5);
                        sg_focus = Some(((l + (1.0 - r)) / 2.0, (t + (1.0 - b)) / 2.0));
                    }
                    // The self-closing form -- `<a:tileRect l=".." t=".."/>` and
                    // the bare `<a:tileRect/>` both arrive here, not above.
                    "tileRect" if sg_in && sg_focus.is_some() && s_gradtile => {
                        let l = get_attr(&e, "l").and_then(|v| gradient_frac(&v)).unwrap_or(0.0);
                        let t = get_attr(&e, "t").and_then(|v| gradient_frac(&v)).unwrap_or(0.0);
                        let r = get_attr(&e, "r").and_then(|v| gradient_frac(&v)).unwrap_or(0.0);
                        let b = get_attr(&e, "b").and_then(|v| gradient_frac(&v)).unwrap_or(0.0);
                        sg_focus = Some(((l + (1.0 - r)) / 2.0, (t + (1.0 - b)) / 2.0));
                    }
                    // Self-closing forms of the cell properties. quick-xml
                    // routes <a:alpha val=".."/> and <a:srgbClr val=".."/> to
                    // Event::Empty, so the Start-arm handlers above see NONE of
                    // them -- the trap this file has hit before.
                    "tcPr" if in_tbl_cell => {
                        // An empty <a:tcPr .../> carries only its attributes.
                        tc_borders = Default::default();
                        tc_fill = None;
                        tc_fill_alpha = None;
                        tc_ln_side = None;
                        tc_mar = (
                            get_attr(&e, "marL").and_then(|v| v.parse::<f32>().ok()).map(emu_to_pt).unwrap_or_else(default_l_ins),
                            get_attr(&e, "marR").and_then(|v| v.parse::<f32>().ok()).map(emu_to_pt).unwrap_or_else(default_r_ins),
                            get_attr(&e, "marT").and_then(|v| v.parse::<f32>().ok()).map(emu_to_pt).unwrap_or_else(default_t_ins),
                            get_attr(&e, "marB").and_then(|v| v.parse::<f32>().ok()).map(emu_to_pt).unwrap_or_else(default_b_ins),
                        );
                        tc_anchor = get_attr(&e, "anchor");
                    }
                    "srgbClr" | "schemeClr" if in_tc_pr => {
                        if let Some(val) = get_attr(&e, "val") {
                            let hex = if name == "srgbClr" {
                                val
                            } else {
                                theme_colors
                                    .get(&val)
                                    .cloned()
                                    .unwrap_or_else(|| scheme_color_to_hex(&val))
                            };
                            if tc_ln_side.is_some() {
                                tc_ln_color = Some(hex);
                            } else {
                                tc_fill = Some(hex);
                            }
                        }
                    }
                    // A run's own colour alpha (d35's 26.9% white numerals).
                    "alpha" if in_outer_shdw => {
                        if let Some(v) = get_attr(&e, "val").and_then(|v| v.parse::<f32>().ok()) {
                            shape_shadow_alpha = (v / 100000.0).clamp(0.0, 1.0);
                        }
                    }
                    "alpha" if in_ln && in_sp_pr => {
                        // The BORDER's translucency (a:ln/a:solidFill/a:alpha).
                        // d49's site pill is a 3pt white ring at 35.3% over
                        // black; opaque it reads as a chalk outline.
                        if let Some(v) = get_attr(&e, "val").and_then(|v| v.parse::<f32>().ok()) {
                            shape_border_alpha = Some((v / 100000.0).clamp(0.0, 1.0));
                        }
                    }
                    "alpha" if in_run && !in_highlight => {
                        if let Some(v) = get_attr(&e, "val") {
                            if let Ok(pc) = v.parse::<f32>() {
                                run_color_alpha = Some((pc / 100000.0).clamp(0.0, 1.0));
                            }
                        }
                    }
                    "alpha" if in_tc_pr => {
                        if let Some(v) = get_attr(&e, "val") {
                            if let Ok(pv) = v.parse::<f32>() {
                                let a = (pv / 100000.0).clamp(0.0, 1.0);
                                if tc_ln_side.is_some() {
                                    tc_ln_alpha = a;
                                } else {
                                    tc_fill_alpha = Some(a);
                                }
                            }
                        }
                    }
                    "tc" if in_tbl_row => {
                        // Self-closing cell — an empty cell.
                        tbl_cur_row.push(TableCell {
                            paragraphs: Vec::new(),
                            fill_color: None,
                            fill_alpha: None,
                            borders: Default::default(),
                            mar_l: default_l_ins(),
                            mar_r: default_r_ins(),
                            mar_t: default_t_ins(),
                            mar_b: default_b_ins(),
                            anchor: None,
                            grid_span: get_attr(&e, "gridSpan")
                                .and_then(|v| v.parse::<u32>().ok())
                                .unwrap_or(1)
                                .max(1),
                            h_merge: matches!(
                                get_attr(&e, "hMerge").as_deref(),
                                Some("1") | Some("true")
                            ),
                        });
                    }
                    "ln" if in_sp_pr => {
                        // Empty <a:ln/> — no border content, just attributes
                        shape_border_dash = None;
                        shape_head_end = None;
                        shape_tail_end = None;
                        shape_line_cap = get_attr(&e, "cap");
                        if let Some(w) = get_attr(&e, "w") {
                            if let Ok(v) = w.parse::<f32>() {
                                shape_border_width = Some(v / 12700.0);
                            }
                        }
                    }
                    // `<a:prstDash val="dash"/>` is ALWAYS self-closing, so this
                    // Empty arm is the one that actually fires; the Start arm
                    // above exists only for symmetry.
                    "prstDash" if in_ln => {
                        shape_border_dash = get_attr(&e, "val").filter(|v| v != "solid");
                    }
                    // Same for the two line ends -- attribute-only elements.
                    "headEnd" | "tailEnd" if in_ln => {
                        let end = parse_line_end(&e);
                        if name == "headEnd" {
                            shape_head_end = end;
                        } else {
                            shape_tail_end = end;
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
                    "gd" if in_prst_geom => {
                        if let (Some(name), Some(fmla)) = (get_attr(&e, "name"), get_attr(&e, "fmla")) {
                            if let Some(value) = fmla
                                .strip_prefix("val ")
                                .and_then(|v| v.trim().parse::<f32>().ok())
                            {
                                shape_adjustments.insert(name, value);
                            }
                        }
                    }
                    // Every custGeom point is self-closing, so the path data
                    // arrives here and NOT in the Start arm; only the commands
                    // that wrap them are Start events.
                    "pt" if cg_pending.is_some() => {
                        if let Some((_, pts)) = cg_pending.as_mut() {
                            let x = get_attr(&e, "x").and_then(|v| v.parse::<f32>().ok());
                            let y = get_attr(&e, "y").and_then(|v| v.parse::<f32>().ok());
                            match (x, y) {
                                (Some(x), Some(y)) => pts.push((x, y)),
                                _ => cg_unsupported = true,
                            }
                        }
                    }
                    "close" if in_cust_geom => {
                        if let Some(p) = cg_cur.as_mut() {
                            p.commands.push(GeomCmd::Close);
                        }
                    }
                    "arcTo" | "quadBezTo" if in_cust_geom => {
                        cg_unsupported = true;
                    }
                    "ph" if in_shape => {
                        shape_ph_type = match get_attr(&e, "type") {
                            Some(t) if !t.is_empty() => Some(t),
                            _ => Some("obj".to_string()),
                        };
                        shape_ph_idx = get_attr(&e, "idx");
                    }
                    "off" if in_grp_xfrm => {
                        if let Some(x) = get_attr(&e, "x") {
                            g_off.0 = x.parse::<f32>().map(emu_to_pt).unwrap_or(0.0);
                        }
                        if let Some(y) = get_attr(&e, "y") {
                            g_off.1 = y.parse::<f32>().map(emu_to_pt).unwrap_or(0.0);
                        }
                    }
                    "chOff" if in_grp_xfrm => {
                        if let Some(x) = get_attr(&e, "x") {
                            g_ch_off.0 = x.parse::<f32>().map(emu_to_pt).unwrap_or(0.0);
                        }
                        if let Some(y) = get_attr(&e, "y") {
                            g_ch_off.1 = y.parse::<f32>().map(emu_to_pt).unwrap_or(0.0);
                        }
                    }
                    "chExt" if in_grp_xfrm => {
                        if let Some(cx) = get_attr(&e, "cx") {
                            g_ch_ext.0 = cx.parse::<f32>().map(emu_to_pt).unwrap_or(0.0);
                        }
                        if let Some(cy) = get_attr(&e, "cy") {
                            g_ch_ext.1 = cy.parse::<f32>().map(emu_to_pt).unwrap_or(0.0);
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
                    "ext" if in_grp_xfrm => {
                        if let Some(cx) = get_attr(&e, "cx") {
                            g_ext.0 = cx.parse::<f32>().map(emu_to_pt).unwrap_or(0.0);
                        }
                        if let Some(cy) = get_attr(&e, "cy") {
                            g_ext.1 = cy.parse::<f32>().map(emu_to_pt).unwrap_or(0.0);
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
                    // <a:alphaModFix amt=".."/> is self-closing, so it arrives
                    // in the Empty arm; both arms carry the handler because a
                    // producer may write either form.
                    "alphaModFix" if in_shape && shape_is_image => {
                        if let Some(v) = get_attr(&e, "amt") {
                            if let Ok(a) = v.parse::<f32>() {
                                shape_image_alpha = Some((a / 100000.0).clamp(0.0, 1.0));
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
                    "srcRect" if in_shape && shape_is_image => {
                        // a:srcRect is SELF-CLOSING (<a:srcRect .../>) — it
                        // arrives as Event::Empty, never Start (the chart
                        // r:id / c:legend / autoTitleDeleted trap).
                        let pct = |v: Option<String>| -> f32 {
                            v.and_then(|s| s.parse::<f32>().ok())
                                .map(|x| x / 100000.0)
                                .unwrap_or(0.0)
                        };
                        shape_src_rect = Some((
                            pct(get_attr(&e, "l")),
                            pct(get_attr(&e, "t")),
                            pct(get_attr(&e, "r")),
                            pct(get_attr(&e, "b")),
                        ));
                    }
                    "fillRect" if in_shape && shape_is_image => {
                        // a:fillRect inside a:stretch is SELF-CLOSING too.
                        let pct = |v: Option<String>| -> f32 {
                            v.and_then(|s| s.parse::<f32>().ok())
                                .map(|x| x / 100000.0)
                                .unwrap_or(0.0)
                        };
                        shape_fill_rect = Some((
                            pct(get_attr(&e, "l")),
                            pct(get_attr(&e, "t")),
                            pct(get_attr(&e, "r")),
                            pct(get_attr(&e, "b")),
                        ));
                    }
                    "rPr" if in_run => {
                        if let Some(b) = get_attr(&e, "b") {
                            run_bold = Some(b == "1" || b == "true");
                        }
                        if let Some(i) = get_attr(&e, "i") {
                            run_italic = i == "1" || i == "true";
                        }
                        if let Some(u) = get_attr(&e, "u") {
                            run_underline = u != "none";
                        }
                        if let Some(sz) = get_attr(&e, "sz") {
                            if let Ok(v) = sz.parse::<f32>() {
                                run_font_size = Some(v / 100.0);
                            }
                        }
                        if s_letterspc {
                            if let Some(spc) = get_attr(&e, "spc") {
                                // Hundredths of a point, and may be negative.
                                if let Ok(v) = spc.parse::<f32>() {
                                    run_spacing = Some(v / 100.0);
                                }
                            }
                        }
                    }
                    // `<a:prstTxWarp prst="textPlain"/>` is self-closing, so
                    // THIS is the arm that fires; the Start arm above exists
                    // only for the rare form with children.
                    "prstTxWarp" if in_body_pr => {
                        shape_text_warp = get_attr(&e, "prst");
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
                        // a:bodyPr/@wrap="none" — the text runs past the box
                        // instead of breaking. The `embedsplit` probes are
                        // COM-built and every arm's box is auto-sized to
                        // 14.5pt around 20pt text, so a renderer that wraps
                        // them anyway shreds "AAABBB" into six lines where
                        // PowerPoint draws one.
                        if let Some(w) = get_attr(&e, "wrap") {
                            shape_wrap = w != "none";
                        }
                        // `a:bodyPr/@spcFirstLastPara` -- keep the first
                        // paragraph's spcBef instead of dropping it. See
                        // `Shape::spc_first_last_para` for the probe that
                        // settled what it does.
                        if let Some(v) = get_attr(&e, "spcFirstLastPara") {
                            shape_spc_first_last = v == "1" || v == "true";
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
                    "spcPts" if in_ln_spc => {
                        // a:lnSpc/a:spcPts/@val is an EXACT line height in
                        // 100ths of a point, not a multiple. 1030 of them over
                        // 3 blind decks; none in the dev corpus.
                        if s_lnspcpts {
                            if let Some(v) = get_attr(&e, "val") {
                                if let Ok(x) = v.parse::<f32>() {
                                    para_line_spacing_pts = Some(x / 100.0);
                                }
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
                            if in_outer_shdw {
                                shape_shadow_color = Some(val);
                            } else if in_bg_pr {
                                slide_background_color = Some(val);
                            } else if in_ln && in_sp_pr {
                                shape_border_color = Some(val);
                            } else if in_sp_pr && !in_ln && !(in_effect_lst && s_effectclr) {
                                shape_fill_color = Some(val);
                            } else if in_run && in_highlight && s_highlight {
                                run_highlight = Some(val);
                            } else if in_run {
                                run_color = Some(val);
                            }
                        }
                    }
                    "schemeClr" => {
                        if let Some(val) = get_attr(&e, "val") {
                            let hex = theme_colors.get(&val).cloned().unwrap_or_else(|| scheme_color_to_hex(&val));
                            if in_outer_shdw {
                                shape_shadow_color = Some(hex);
                            } else if in_bg_pr {
                                slide_background_color = Some(hex);
                            } else if in_ln && in_sp_pr {
                                shape_border_color = Some(hex);
                            } else if in_sp_pr && !in_ln && !(in_effect_lst && s_effectclr) {
                                shape_fill_color = Some(hex);
                            } else if in_run && in_highlight && s_highlight {
                                run_highlight = Some(hex);
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
                    "effectLst" if in_effect_lst => {
                        in_effect_lst = false;
                    }
                    "outerShdw" if in_outer_shdw => {
                        in_outer_shdw = false;
                        if let Some((blur, dist, dir)) = shape_shadow_draft.take() {
                            shape_shadow = Some(crate::ir::ShapeShadow {
                                blur_pt: blur,
                                dist_pt: dist,
                                dir_deg: dir,
                                // The schema's default shadow colour is black.
                                color: shape_shadow_color
                                    .take()
                                    .unwrap_or_else(|| "000000".to_string()),
                                alpha: shape_shadow_alpha,
                            });
                        }
                    }
                    "highlight" if in_highlight => {
                        in_highlight = false;
                    }
                    "gs" if sg_in_gs => {
                        sg_in_gs = false;
                        if let Some(c) = sg_color.take() {
                            sg_stops.push(SlideGradientStop {
                                pos: sg_pos,
                                color: c,
                                alpha: sg_alpha,
                            });
                        }
                    }
                    "path" if sg_in_path => sg_in_path = false,
                    "gradFill" if sg_in => sg_in = false,
                    "prstGeom" if in_prst_geom => {
                        in_prst_geom = false;
                    }
                    "moveTo" | "lnTo" | "cubicBezTo" if in_cust_geom => {
                        // The points are complete only at the end tag. A command
                        // with the wrong arity is a shape we do not understand,
                        // so it poisons the whole geometry rather than silently
                        // dropping one segment.
                        if let Some((kind, pts)) = cg_pending.take() {
                            match (kind, pts.as_slice()) {
                                ("moveTo", [(x, y)]) => {
                                    if let Some(p) = cg_cur.as_mut() {
                                        p.commands.push(GeomCmd::MoveTo(*x, *y));
                                    }
                                }
                                ("lnTo", [(x, y)]) => {
                                    if let Some(p) = cg_cur.as_mut() {
                                        p.commands.push(GeomCmd::LineTo(*x, *y));
                                    }
                                }
                                ("cubicBezTo", [(x1, y1), (x2, y2), (x3, y3)]) => {
                                    if let Some(p) = cg_cur.as_mut() {
                                        p.commands
                                            .push(GeomCmd::CubicTo(*x1, *y1, *x2, *y2, *x3, *y3));
                                    }
                                }
                                _ => cg_unsupported = true,
                            }
                        }
                    }
                    "path" if in_cust_geom => {
                        if let Some(p) = cg_cur.take() {
                            cg_paths.push(p);
                        }
                    }
                    "custGeom" if in_cust_geom => {
                        in_cust_geom = false;
                        cg_pending = None;
                    }
                    "solidFill" if in_solid_fill => {
                        in_solid_fill = false;
                    }
                    "xfrm" if in_grp_xfrm => {
                        in_grp_xfrm = false;
                        // Compose with the enclosing group's transform (this
                        // group's own off/ext live in the PARENT's child space).
                        let base = if grp_stack.len() >= 2 {
                            grp_stack[grp_stack.len() - 2]
                        } else {
                            (0.0, 0.0, 1.0, 1.0, 0.0, 0.0, 0.0, false, false)
                        };
                        let mut box_x = base.0 + g_off.0 * base.2;
                        let mut box_y = base.1 + g_off.1 * base.3;
                        let box_w = g_ext.0 * base.2;
                        let box_h = g_ext.1 * base.3;
                        // ★A nested group's own BOX has to turn and mirror with
                        // its parent, exactly as a leaf shape does. Only the
                        // leaf did: the flip and the rotation were carried
                        // forward in the tuple (so the leaf mirrored its
                        // geometry) while the nested group's box stayed where
                        // the un-flipped parent would have put it. Shapes
                        // DIRECTLY in a flipped group therefore landed right and
                        // shapes one level deeper landed wrong -- d10 s11's
                        // pizza, whose slice is a direct child and whose
                        // pepperoni highlights are a level down, came out with
                        // every pepperoni wearing a tail.
                        if s_grpnest && (base.7 || base.8 || base.4.abs() > 1e-4) {
                            let (mut t, mut u) =
                                (box_x + box_w / 2.0 - base.5, box_y + box_h / 2.0 - base.6);
                            if base.7 {
                                t = -t;
                            }
                            if base.8 {
                                u = -u;
                            }
                            let (sn, cs) = (base.4.to_radians().sin(), base.4.to_radians().cos());
                            box_x = base.5 + t * cs - u * sn - box_w / 2.0;
                            box_y = base.6 + t * sn + u * cs - box_h / 2.0;
                        }
                        let sx = if g_ch_ext.0 != 0.0 { box_w / g_ch_ext.0 } else { base.2 };
                        let sy = if g_ch_ext.1 != 0.0 { box_h / g_ch_ext.1 } else { base.3 };
                        if let Some(top) = grp_stack.last_mut() {
                            *top = (
                                box_x - g_ch_off.0 * sx,
                                box_y - g_ch_off.1 * sy,
                                sx,
                                sy,
                                base.4 + g_rot,
                                box_x + box_w / 2.0,
                                box_y + box_h / 2.0,
                                base.7 ^ (g_flip.0 && s_grpflip),
                                base.8 ^ (g_flip.1 && s_grpflip),
                            );
                        }
                    }
                    "grpSpPr" if in_grp_sp_pr => {
                        in_grp_sp_pr = false;
                    }
                    "grpSp" if !grp_stack.is_empty() => {
                        grp_stack.pop();
                    }
                    "sp" | "pic" | "cxnSp" if in_shape => {
                        // A connector is a shape here but not one the editor
                        // counts, so it takes no index and does not advance the
                        // count.
                        let editable = name != "cxnSp";
                        let this_sp_index = editable.then_some(sp_count);
                        if editable {
                            sp_count += 1;
                        }
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
                        // A shape inside a p:grpSp expresses its geometry in
                        // the group's child space.  A placeholder that fell
                        // back to layout/master geometry is already in slide
                        // space, so it is left alone.
                        let mut extra_rot = 0.0f32;
                        let mut grp_flip = (false, false);
                        let (use_x, use_y, use_w, use_h) = match grp_stack.last() {
                            Some(&(ox, oy, sx, sy, rot, cx, cy, fh, fv))
                                if s_grp && shape_has_xfrm =>
                            {
                                let (x, y, w, h) =
                                    (ox + use_x * sx, oy + use_y * sy, use_w * sx, use_h * sy);
                                if (rot.abs() > 1e-4 || fh || fv) && s_grprot {
                                    // The group turns and mirrors as a whole
                                    // about its own centre: the child's CENTRE
                                    // moves with it, the child keeps its size,
                                    // and its own turn and mirror compose with
                                    // the group's. OOXML mirrors first, then
                                    // rotates.
                                    let (mut t, mut u) = (x + w / 2.0 - cx, y + h / 2.0 - cy);
                                    if fh {
                                        t = -t;
                                    }
                                    if fv {
                                        u = -u;
                                    }
                                    let (s, c) =
                                        (rot.to_radians().sin(), rot.to_radians().cos());
                                    let (nx, ny) = (cx + t * c - u * s, cy + t * s + u * c);
                                    extra_rot = rot;
                                    grp_flip = (fh, fv);
                                    (nx - w / 2.0, ny - h / 2.0, w, h)
                                } else {
                                    (x, y, w, h)
                                }
                            }
                            _ => (use_x, use_y, use_w, use_h),
                        };
                        // A single-axis mirror reverses the child's own turn.
                        let own_rot = if grp_flip.0 ^ grp_flip.1 {
                            -shape_rotation
                        } else {
                            shape_rotation
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
                            sp_index: this_sp_index,
                            top_level: grp_stack.is_empty(),
                            x: use_x,
                            y: use_y,
                            width: use_w,
                            height: use_h,
                            rotation: own_rot + extra_rot,
                            flip_h: shape_flip_h ^ grp_flip.0,
                            flip_v: shape_flip_v ^ grp_flip.1,
                            shape_type: shape_prst.take(),
                            adjustments: std::mem::take(&mut shape_adjustments),
                            ph_type: shape_ph_type.clone(),
                            ph_levels: merge_ph_levels(
                                lookup_ph_levels(
                                    &layout_ph_styles,
                                    shape_ph_type.as_ref(),
                                    shape_ph_idx.as_ref(),
                                ),
                                lookup_ph_levels_in(
                                    &master_ph_styles,
                                    shape_ph_type.take().as_ref(),
                                    shape_ph_idx.as_ref(),
                                    true,
                                ),
                            ),
                            content,
                            fill_color: shape_fill_color.take(),
                            fill_alpha: shape_fill_alpha.take(),
                            border_color: shape_border_color.take(),
                            border_alpha: shape_border_alpha.take(),
                            border_width: shape_border_width.take(),
                            border_dash: shape_border_dash.take(),
                            head_end: shape_head_end.take(),
                            tail_end: shape_tail_end.take(),
                            line_cap: shape_line_cap.take(),
                            l_ins: shape_l_ins,
                            r_ins: shape_r_ins,
                            t_ins: shape_t_ins,
                            b_ins: shape_b_ins,
                            anchor: resolved_anchor,
                            wrap_text: std::mem::replace(&mut shape_wrap, true),
                            spc_first_last_para: std::mem::replace(
                                &mut shape_spc_first_last,
                                false,
                            ),
                            text_warp: shape_text_warp.take(),
                            src_rect: shape_src_rect.take(),
                            fill_rect: shape_fill_rect.take(),
                            rot_with_shape: shape_rot_with_shape,
                            image_alpha: shape_image_alpha,
                            gradient: if sg_stops.len() >= 2 {
                                Some(SlideGradient {
                                    stops: std::mem::take(&mut sg_stops),
                                    angle_deg: sg_angle,
                                    scaled: sg_scaled,
                                    focus: sg_focus,
                                    rot_with_shape: sg_rot_with_shape,
                                })
                            } else {
                                sg_stops.clear();
                                None
                            },
                            custom_geometry: take_custom_geometry(
                                &mut cg_paths,
                                &mut cg_unsupported,
                            ),
                            shadow: shape_shadow.take(),
                        });
                        in_shape = false;
                    }
                    "p" if in_paragraph => {
                        in_paragraph = false;
                        let para = SlideParagraph {
                            runs: std::mem::take(&mut para_runs),
                            alignment: para_alignment,
                            line_spacing: para_line_spacing,
                            line_spacing_pts: para_line_spacing_pts,
                            space_before: para_space_before,
                            space_after: para_space_after,
                            lvl: para_lvl,
                            mar_l: para_mar_l,
                            indent: para_indent,
                            bullet: std::mem::take(&mut para_bullet),
                            end_para_size: para_end_size.take(),
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
                    "lnL" | "lnR" | "lnT" | "lnB" if in_tc_pr => {
                        if let (Some(side), Some(color)) = (tc_ln_side.take(), tc_ln_color.take()) {
                            tc_borders[side] = Some(CellBorder {
                                color,
                                width: tc_ln_width,
                                alpha: tc_ln_alpha,
                            });
                        }
                        tc_ln_side = None;
                    }
                    "tcPr" if in_tc_pr => {
                        in_tc_pr = false;
                    }
                    "tc" if in_tbl_cell => {
                        in_tbl_cell = false;
                        tbl_cur_row.push(TableCell {
                            paragraphs: std::mem::take(&mut tbl_cur_cell_paragraphs),
                            fill_color: tc_fill.take(),
                            fill_alpha: tc_fill_alpha.take(),
                            borders: std::mem::take(&mut tc_borders),
                            mar_l: tc_mar.0,
                            mar_r: tc_mar.1,
                            mar_t: tc_mar.2,
                            mar_b: tc_mar.3,
                            anchor: tc_anchor.take(),
                            grid_span: std::mem::replace(&mut tc_grid_span, 1),
                            h_merge: std::mem::take(&mut tc_h_merge),
                        });
                        tc_mar = (default_l_ins(), default_r_ins(), default_t_ins(), default_b_ins());
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
                            // A graphicFrame is not one of the elements the
                            // editor counts, so it has no index in its numbering.
                            sp_index: None,
                            top_level: true,
                            wrap_text: true,
                            spc_first_last_para: false,
                            text_warp: None,
                            x: shape_x,
                            y: shape_y,
                            width: shape_w,
                            height: shape_h,
                            rotation: shape_rotation,
                            flip_h: shape_flip_h,
                            flip_v: shape_flip_v,
                            shape_type: shape_prst.take(),
                            adjustments: std::mem::take(&mut shape_adjustments),
                            ph_type: None,
                            gradient: None,
                            ph_levels: Vec::new(),
                            content,
                            fill_color: shape_fill_color.take(),
                            fill_alpha: shape_fill_alpha.take(),
                            border_color: shape_border_color.take(),
                            border_alpha: shape_border_alpha.take(),
                            border_width: shape_border_width.take(),
                            border_dash: shape_border_dash.take(),
                            head_end: shape_head_end.take(),
                            tail_end: shape_tail_end.take(),
                            line_cap: shape_line_cap.take(),
                            l_ins: shape_l_ins,
                            r_ins: shape_r_ins,
                            t_ins: shape_t_ins,
                            b_ins: shape_b_ins,
                            anchor: None,
                            src_rect: shape_src_rect.take(),
                            fill_rect: shape_fill_rect.take(),
                            rot_with_shape: shape_rot_with_shape,
                            image_alpha: shape_image_alpha,
                            custom_geometry: take_custom_geometry(
                                &mut cg_paths,
                                &mut cg_unsupported,
                            ),
                            shadow: shape_shadow.take(),
                        });
                    }
                    "r" | "fld" if in_run => {
                        in_run = false;
                        // A field's own text is what PowerPoint computes, not
                        // what the file caches. `slidenum` is the only type in
                        // the dev corpus (295 fields over 9 decks) and prints
                        // the slide's 1-based position offset by the deck's
                        // firstSlideNum; any other type keeps its cached text,
                        // which is the last value PowerPoint itself wrote.
                        if let Some(kind) = fld_type.take() {
                            if kind == "slidenum" {
                                if s_slidenum {
                                    run_text = (slide_index as u32 + first_slide_num - 1)
                                        .to_string();
                                } else {
                                    run_text.clear();
                                }
                            } else if !s_slidenum {
                                run_text.clear();
                            }
                        }
                        if !run_text.is_empty() {
                            para_runs.push(SlideRun {
                                text: std::mem::take(&mut run_text),
                                font_size: run_font_size,
                                bold: run_bold,
                                italic: run_italic,
                            underline: run_underline,
                                color: run_color.take(),
                            color_alpha: run_color_alpha.take(),
                                highlight: run_highlight.take(),
                                font_family: run_font_family.take(),
                                spacing: run_spacing,
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

    // PowerPoint's draw order is background, master, layout, slide, and the
    // renderer walks `shapes` in order, so PREPENDING the inherited ones (the
    // master's first, as collected) reproduces it with no renderer change.
    let shapes = if inherited_shapes.is_empty() {
        shapes
    } else {
        inherited_shapes.extend(shapes);
        inherited_shapes
    };

    Ok(Slide {
        index: slide_index,
        shapes,
        background_color: if slide_has_own_bg {
            slide_background_color
        } else {
            slide_background_color.or(inherited_bg)
        },
        background_gradient: if std::env::var("OXI_BGGRAD_DISABLE").is_ok() {
            None
        } else if slide_has_own_bg {
            // Same ownership rule as the flat colour: a slide that declares any
            // <p:bg> of its own never reaches back to the layout/master.
            parse_bg_gradient(xml, theme_colors)
        } else {
            inherited_grad
        },
        // Picture background. Same ownership rule again, and the same env gate
        // shape as the gradient so the whole feature can be switched off for an
        // A/B without touching the other two fills.
        background_image: if !bgimg_on {
            None
        } else if slide_has_own_bg {
            load_bg_image(xml, slide_rels_path, archive)
        } else {
            inherited_img
        },
    })
}

/// `a:fillToRect` / `a:tileRect` edge value -> fraction of the page.
///
/// Both notations occur in the corpus: `"50%"` (13 of the 52 background
/// gradients) and the DrawingML integer 1000ths, e.g. `"100000"` (39).
fn gradient_frac(v: &str) -> Option<f32> {
    let v = v.trim();
    if let Some(p) = v.strip_suffix('%') {
        p.trim().parse::<f32>().ok().map(|x| x / 100.0)
    } else {
        v.parse::<f32>().ok().map(|x| x / 100_000.0)
    }
}

/// Extract a GRADIENT background from a slide / layout / master part.
///
/// `p:cSld/p:bg/p:bgPr/a:gradFill` -> the `a:gsLst` stops plus either `a:lin`
/// (an angle) or `a:path path="circle"` (a focus from `a:fillToRect`).
///
/// PowerPoint render-truth. PowerPoint exports a gradient background as a PDF
/// *Pattern* (`/Pattern cs /Pn scn` over a full-page rectangle) whose Shading
/// is axial (type 2) for `a:lin` and radial (type 3) for `a:path`, so the
/// geometry is readable exactly:
///   * `a:lin` ang=0 runs left->right, 90 top->bottom, 270 bottom->top. The
///     axis is centred on the page and spans it: probe B1's axis is
///     (0,270)->(720,270), i.e. the full page width.
///   * `a:path path="circle"` puts t=0 at the point `a:fillToRect` describes
///     and t=1 at the FARTHEST page corner. Measured twice on the dev corpus:
///     d04 (`50%` on all four edges) gives centre (360,202.5) with r=413.05,
///     the distance to every corner of a 720x405 page; d15 (`l=t=100000`,
///     i.e. the bottom-right corner) gives centre (720,0) -- PDF y is up --
///     with r=826.09, the distance to the opposite corner.
///   * `path="rect"` is the one form PowerPoint rasterizes instead of emitting
///     a shading, and no corpus background uses it, so it is left unpainted
///     rather than guessed at.
///   * An EMPTY `a:tileRect` overrides `a:fillToRect` and centres the ramp
///     (probe arms B7/B8/B9 all render the same centred circle). The corpus
///     pairs every bottom-right focus with a tileRect that carries attributes
///     and every centred focus with an empty one, so the two readings agree on
///     every corpus slide. No background gradFill carries a colour modifier,
///     so the bare stop colour is exact.
fn parse_bg_gradient(xml: &str, theme_colors: &HashMap<String, String>) -> Option<SlideGradient> {
    let mut reader = Reader::from_str(xml);
    let mut in_bg = false;
    let mut in_bg_pr = false;
    let mut in_grad = false;
    let mut in_gs = false;
    let mut in_path = false;
    let mut cur_pos: f32 = 0.0;
    let mut cur_color: Option<String> = None;
    let mut cur_alpha: f32 = 1.0;
    let mut stops: Vec<SlideGradientStop> = Vec::new();
    let mut angle_deg: Option<f32> = None;
    let mut scaled = false;
    let mut focus: Option<(f32, f32)> = None;
    let mut empty_tile_rect = false;
    // An absent `a:fillToRect` inset is 0; `OXI_GRADFTR_DISABLE` restores the
    // old 0.5, which only differs when some edges are given and others are not.
    let ftr_default: f32 = if std::env::var("OXI_GRADFTR_DISABLE").is_err() { 0.0 } else { 0.5 };
    loop {
        let ev = match reader.read_event() {
            Ok(ev) => ev,
            Err(_) => break,
        };
        match ev {
            Event::Start(e) | Event::Empty(e) => {
                let name = local_name(e.name().as_ref());
                match name.as_str() {
                    "bg" => in_bg = true,
                    "bgPr" if in_bg => in_bg_pr = true,
                    "gradFill" if in_bg_pr => in_grad = true,
                    "gs" if in_grad => {
                        in_gs = true;
                        cur_pos = get_attr(&e, "pos")
                            .and_then(|v| gradient_frac(&v))
                            .unwrap_or(0.0);
                        cur_color = None;
                        cur_alpha = 1.0;
                    }
                    "alpha" if in_gs => {
                        if let Some(v) = get_attr(&e, "val") {
                            if let Ok(a) = v.parse::<f32>() {
                                cur_alpha = (a / 100000.0).clamp(0.0, 1.0);
                            }
                        }
                    }
                    "srgbClr" if in_gs && cur_color.is_none() => {
                        cur_color = get_attr(&e, "val");
                    }
                    "schemeClr" if in_gs && cur_color.is_none() => {
                        if let Some(v) = get_attr(&e, "val") {
                            cur_color = Some(
                                theme_colors
                                    .get(&v)
                                    .cloned()
                                    .unwrap_or_else(|| scheme_color_to_hex(&v)),
                            );
                        }
                    }
                    "lin" if in_grad => {
                        angle_deg = get_attr(&e, "ang")
                            .and_then(|v| v.parse::<f32>().ok())
                            .map(|v| v / 60_000.0);
                        scaled = get_attr(&e, "scaled").as_deref() == Some("1");
                    }
                    "path" if in_grad => {
                        if get_attr(&e, "path").as_deref() == Some("circle") {
                            in_path = true;
                            focus = Some((0.5, 0.5));
                        }
                    }
                    "fillToRect" if in_path => {
                        // S-GRADFTR (2026-08-27). An ABSENT inset is 0, not
                        // 0.5. Blind d15 writes `<a:fillToRect b="100%"
                        // r="100%"/>` with l and t absent; defaulting those to
                        // 0.5 put the focus at (0.25,0.25) instead of the
                        // top-left corner. PowerPoint's own PDF settles it --
                        // the slide's radial shading is
                        //     /Matrix [826.09 0 0 826.09 0 405]
                        //     /Coords [0 0 0  0 0 1]
                        // i.e. centre (0,405) = the TOP-LEFT corner of a 720x405
                        // page (PDF y is up) and radius 826.09 = sqrt(720^2 +
                        // 405^2), the distance to the far corner. With the focus
                        // at (0,0) the ramp Oxi already implements predicts that
                        // page's background to within ~2/255; with (0.25,0.25)
                        // it ran up to 30/255 dark.
                        let d = ftr_default;
                        let l = get_attr(&e, "l").and_then(|v| gradient_frac(&v)).unwrap_or(d);
                        let t = get_attr(&e, "t").and_then(|v| gradient_frac(&v)).unwrap_or(d);
                        let r = get_attr(&e, "r").and_then(|v| gradient_frac(&v)).unwrap_or(d);
                        let b = get_attr(&e, "b").and_then(|v| gradient_frac(&v)).unwrap_or(d);
                        // The rect's centre is the focus: (1,1,0,0) collapses to
                        // the bottom-right corner and 50% on all edges to the
                        // page centre.
                        focus = Some(((l + (1.0 - r)) / 2.0, (t + (1.0 - b)) / 2.0));
                    }
                    "tileRect" if in_grad => {
                        empty_tile_rect = e.attributes().next().is_none();
                    }
                    _ => {}
                }
            }
            Event::End(e) => match local_name(e.name().as_ref()).as_str() {
                "gs" => {
                    if let Some(c) = cur_color.take() {
                        stops.push(SlideGradientStop {
                            alpha: cur_alpha,
                            pos: cur_pos,
                            color: c,
                        });
                    }
                    in_gs = false;
                }
                "path" => in_path = false,
                "gradFill" => in_grad = false,
                "bgPr" => in_bg_pr = false,
                "bg" => break,
                _ => {}
            },
            Event::Eof => break,
            _ => {}
        }
    }
    // An EMPTY `<a:tileRect/>` makes PowerPoint ignore `a:fillToRect` and
    // centre the ramp: probe arms B7 (focus bottom-right), B8 (top-left) and
    // B9 (centred) all render as the same centred circle.  The corpus never
    // mixes the two -- all 38 bottom-right focuses carry a tileRect with
    // attributes and all 14 centred ones an empty tileRect -- so this only
    // decides hand-written combinations, and it leaves every corpus focus
    // exactly where reading `fillToRect` alone would put it.
    if empty_tile_rect && focus.is_some() {
        focus = Some((0.5, 0.5));
    }
    if stops.len() < 2 || (angle_deg.is_none() && focus.is_none()) {
        return None;
    }
    stops.sort_by(|a, b| a.pos.partial_cmp(&b.pos).unwrap_or(std::cmp::Ordering::Equal));
    Some(SlideGradient {
        stops,
        angle_deg,
        scaled,
        focus,
        // A page background has no shape transform to ride.
        rot_with_shape: true,
    })
}

/// Extract the SOLID background colour from a slide / layout / master part.
///
/// `p:cSld/p:bg/p:bgPr/a:solidFill` -> `a:srgbClr@val`, or `a:schemeClr@val`
/// resolved through the theme colour scheme (Spec #10).  Returns None for
/// `gradFill` / `blipFill` / `bgRef`, which are not a single colour.
///
/// Dev-corpus measurement (40 decks, 886 slides): 82.8% of effective slide
/// backgrounds are a plain solidFill, and NONE of the 651 solidFill `<p:bg>`
/// blocks carries a colour modifier (lumMod / lumOff / tint / shade / alpha),
/// so reading the bare value is exact.
fn parse_bg_solid_fill(xml: &str, theme_colors: &HashMap<String, String>) -> Option<String> {
    let mut reader = Reader::from_str(xml);
    let mut in_bg = false;
    let mut in_bg_pr = false;
    let mut in_solid = false;
    let mut found: Option<String> = None;
    loop {
        let ev = match reader.read_event() {
            Ok(ev) => ev,
            Err(_) => break,
        };
        match ev {
            Event::Start(e) | Event::Empty(e) => {
                let name = local_name(e.name().as_ref());
                match name.as_str() {
                    "bg" => in_bg = true,
                    "bgPr" if in_bg => in_bg_pr = true,
                    "solidFill" if in_bg_pr => in_solid = true,
                    "srgbClr" if in_solid && found.is_none() => {
                        found = get_attr(&e, "val");
                    }
                    "schemeClr" if in_solid && found.is_none() => {
                        if let Some(v) = get_attr(&e, "val") {
                            found = Some(
                                theme_colors
                                    .get(&v)
                                    .cloned()
                                    .unwrap_or_else(|| scheme_color_to_hex(&v)),
                            );
                        }
                    }
                    _ => {}
                }
            }
            Event::End(e) => match local_name(e.name().as_ref()).as_str() {
                "solidFill" => in_solid = false,
                "bgPr" => in_bg_pr = false,
                // The background is the first thing in p:cSld, so stop as soon
                // as it closes rather than walking the whole part.
                "bg" => break,
                _ => {}
            },
            Event::Eof => break,
            _ => {}
        }
    }
    found
}

/// Resolve a background picture fill declared in `xml` against the rels of the
/// part it came from, and load the media bytes.
///
/// `rels_path` is the `_rels/<part>.rels` of that same part -- a slide, layout
/// and master each carry their own, and the same `rId3` means different images
/// in each, so the caller must pass the matching one.
fn load_bg_image(
    xml: &str,
    rels_path: &str,
    archive: &mut OoxmlArchive,
) -> Option<SlideBackgroundImage> {
    let rid = parse_bg_blip_rid(xml)?;
    let rels_xml = archive.try_read_part(rels_path).ok()??;
    let rels = parse_relationships(&rels_xml).ok()?;
    let rel = rels.get(&rid)?;
    let image_path = resolve_slide_relative_path(rels_path, &rel.target);
    let data = archive.read_binary_part(&image_path).ok()?;
    if data.is_empty() {
        return None;
    }
    Some(SlideBackgroundImage {
        data,
        content_type: detect_content_type(&rel.target),
    })
}

/// Does this slide / layout part switch OFF the master's own shapes?
///
/// OOXML gates a master's non-placeholder shapes behind `p:sld/@showMasterSp`
/// and `p:sldLayout/@showMasterSp`, both defaulting to "1". Measured on the dev
/// corpus the attribute is never 0, but honouring it costs one attribute read
/// and a deck that sets it would otherwise gain ink PowerPoint does not draw.
fn show_master_sp_off(xml: &str) -> bool {
    let mut reader = Reader::from_str(xml);
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) | Ok(Event::Empty(ref e)) => {
                let name = local_name(e.name().as_ref());
                if name == "sld" || name == "sldLayout" {
                    return get_attr(e, "showMasterSp").as_deref() == Some("0");
                }
            }
            Ok(Event::Eof) | Err(_) => return false,
            _ => {}
        }
        buf.clear();
    }
}

/// Does this image's container declare per-pixel transparency?
///
/// This crate has no image decoder, so the test is on container bytes: for PNG
/// the IHDR colour type (4 = grey+alpha, 6 = RGBA) or a `tRNS` chunk, which the
/// spec requires to precede `IDAT`. JPEG and BMP carry no alpha. Anything else
/// is treated as transparent rather than risk painting it wrong.
///
/// Checked against decoded pixels for every inherited picture in the dev
/// corpus: 41 transparent / 2 opaque, with zero disagreements.
///
/// The renderer composites per-pixel alpha now, so this no longer decides
/// whether a picture may ship -- it only feeds the `OXI_LMPICALPHA_DISABLE`
/// escape hatch, which restores the old reject-transparent behaviour.
fn media_has_alpha(data: &[u8]) -> bool {
    if data.starts_with(&[0x89, b'P', b'N', b'G', 0x0d, 0x0a, 0x1a, 0x0a]) {
        if data.len() < 26 {
            return true;
        }
        if matches!(data[25], 4 | 6) {
            return true;
        }
        let end = data
            .windows(4)
            .position(|w| w == b"IDAT")
            .unwrap_or(data.len());
        return data[..end].windows(4).any(|w| w == b"tRNS");
    }
    if data.starts_with(&[0xff, 0xd8]) || data.starts_with(b"BM") {
        return false;
    }
    true
}

/// The non-placeholder shapes a slideLayout / slideMaster draws for ITSELF.
///
/// The parser reads those parts only for placeholder geometry, the text styles
/// and the background, so every shape a layout or master paints on its own
/// account was dropped. Measured on the dev corpus that costs 70 slides their
/// ink, and 20 of them a FULL-PAGE picture (d19 8, d37 7, d28 5).
///
/// Oxi paints nothing at all from these parts today, so whatever is emitted
/// here can only ADD ink PowerPoint also has -- provided every shape is painted
/// FAITHFULLY. That is what bounds the subset, and the bounds come from the IR
/// and the renderer, not from taste:
///
///   * `fill_color` is a single hex -- no gradient, no colour transform
///   * the generic fill/border paint is AXIS-ALIGNED -- rot/flip are ignored
///   * there are no prstGeom arms, so any preset would be drawn as a RECTANGLE
///
/// So the subset is `p:pic` with a plain opaque `r:embed`, and `p:sp` that is a
/// plain-solidFill `rect`. Everything else is left out rather than painted
/// wrong (the 2ea81a callout lesson: incorrect ink is worse than none).
/// Measured rejects, by reason: 306 custGeom, 186 grpSp, 157 rot, 155 gradFill,
/// 75 outerShdw, 61 cxnSp, 46 translucent pics, 15 ellipse, 11 flip.
///
/// The picture MEDIA used to be judged too, and a transparent PNG rejected,
/// because the blit would have painted it as an opaque slab -- which held back
/// most of the inherited pictures, the full-page ones among them. The renderer
/// composites per-pixel alpha now, so transparency is no longer a reason to
/// drop a picture. `OXI_LMPICALPHA_DISABLE` restores the old rejection.
///
/// A constant `a:alpha` on the shape's own fill is no longer a reason to drop
/// it either: the renderer composites it (S-FILLALPHA), so the two full-page
/// scrims -- d08 layout3 `000000` at 62.01%, d16 layout10 `CFD8DC` at 49.23% --
/// are painted rather than skipped. `OXI_FILLALPHA_DISABLE` restores the old
/// rejection, so that arm still drops them rather than painting them opaque.
/// An alpha ELSEWHERE (a text colour) and the colour TRANSFORMS beside it
/// (lumMod/lumOff/tint/shade/satMod/hueMod) are still rejected, since a flat
/// hex cannot express them.
///
/// `a:alphaModFix` is judged by its `@amt`, not its presence: the attribute
/// defaults to 100000 (fully opaque), and the bare `<a:alphaModFix/>` form --
/// 86 of the 132 layout/master pics, the same degenerate spelling the
/// background census found -- is a no-op. Only a real amt (25000/24000/60000)
/// is translucency.
///
/// Redundant shapes are NOT filtered: an inherited picture that the slide also
/// draws, with the same media and the same box, is simply covered by it, which
/// is exactly what PowerPoint's own layout-then-slide draw order does.
/// Place a box stated in a group's child space: mirror, then turn its CENTRE
/// about the group's own centre, keeping its size. OOXML mirrors before it
/// rotates, and the same operation serves a nested group's box and a leaf
/// shape's.
fn place_in_group(
    x: f32,
    y: f32,
    w: f32,
    h: f32,
    g: &(f32, f32, f32, f32, f32, f32, f32, bool, bool),
) -> (f32, f32) {
    let (_, _, _, _, rot, cx, cy, fh, fv) = *g;
    if rot.abs() <= 1e-4 && !fh && !fv {
        return (x, y);
    }
    let (mut t, mut u) = (x + w / 2.0 - cx, y + h / 2.0 - cy);
    if fh {
        t = -t;
    }
    if fv {
        u = -u;
    }
    let (sn, cs) = (rot.to_radians().sin(), rot.to_radians().cos());
    // Two steps, not one expression: f32 associates differently and a 1-ulp
    // coordinate can flip a pixel, which is enough to break a byte-identity arm.
    let (nx, ny) = (cx + t * cs - u * sn, cy + t * sn + u * cs);
    (nx - w / 2.0, ny - h / 2.0)
}

fn parse_inherited_shapes(
    xml: &str,
    rels_path: &str,
    archive: &mut OoxmlArchive,
    theme_colors: &HashMap<String, String>,
) -> Vec<Shape> {
    let mut reader = Reader::from_str(xml);
    let mut buf = Vec::new();
    let mut out: Vec<Shape> = Vec::new();

    let mut in_tree = false;
    // Nesting depth of shape-kind elements INSIDE p:spTree, so that a grpSp's
    // children are never mistaken for top-level shapes.
    let mut nest: u32 = 0;
    let mut kind: Option<&'static str> = None;

    // Per-candidate state.
    let mut ok = false;
    let mut x = 0.0f32;
    let mut y = 0.0f32;
    let mut w = 0.0f32;
    let mut h = 0.0f32;
    let mut have_off = false;
    let mut have_ext = false;
    let mut prst: Option<String> = None;
    let mut fill: Option<String> = None;
    let mut fill_alpha: Option<f32> = None;
    let mut embed: Option<String> = None;
    let mut src_rect: Option<(f32, f32, f32, f32)> = None;
    let mut fill_rect: Option<(f32, f32, f32, f32)> = None;
    let mut ln_color: Option<String> = None;
    let mut ln_alpha: Option<f32> = None;
    let mut ln_width: Option<f32> = None;
    let mut ln_no_fill = false;
    let mut in_sp_pr = false;
    let mut in_pic_blip = false;
    let mut ln_depth: u32 = 0;
    let mut in_solid = false;
    // <a:outerShdw>: geometry attributes while open, then colour and alpha.
    let mut in_ig_shdw = false;
    let mut ig_shdw_draft: Option<(f32, f32, f32)> = None;
    let mut ig_shdw_color: Option<String> = None;
    let mut ig_shdw_alpha: f32 = 1.0;
    let mut ig_shadow: Option<crate::ir::ShapeShadow> = None;
    // custGeom paths and a gradient fill are drawable now (S-CUSTGEOM /
    // S-SHAPEGRAD), so inherited shapes carrying them are no longer refused.
    let s_inherit_geom = std::env::var("OXI_LMGEOM_DISABLE").is_err();
    let s_gradshape = std::env::var("OXI_GRADALPHA_DISABLE").is_err();
    // An inherited shape keeps its own turn and mirror unless this is set,
    // which restores refusing it outright.
    let s_inherit_rot = std::env::var("OXI_LMROT_DISABLE").is_err();
    // An inherited picture keeps its `a:alphaModFix` opacity unless this is
    // set, which restores refusing the shape.
    let s_inherit_imgalpha = std::env::var("OXI_LMIMGALPHA_DISABLE").is_err();
    let mut ig_img_alpha: Option<f32> = None;
    // An inherited shape whose FILL is a picture is emitted as that picture,
    // clipped to its own geometry, unless this is set.
    let s_inherit_fillimg = std::env::var("OXI_LMFILLIMG_DISABLE").is_err();
    // An inherited shape may keep a preset the renderer can actually draw.
    let s_inherit_prst = std::env::var("OXI_LMPRST_DISABLE").is_err();
    // A layout/master shape can live inside a `p:grpSp`, and this walk treated
    // the whole group as ONE candidate it then refused -- so every shape in it
    // was invisible. 1241 shapes across 452 groups are like that (d19 460,
    // d10 408, d01 64, d03 54, d12 46, d18 42), including the pencils d19
    // slide 1 is built from. Descend instead, carrying the group's child-space
    // mapping the way `parse_slide` does.
    let s_inherit_grp = std::env::var("OXI_LMGRP_DISABLE").is_err();
    // Composing a NESTED group's own placement through its parent's turn and
    // mirror ships after the descent itself, so it needs its own opt-out --
    // disabling both together would compare against the pre-descent build.
    let s_grp_nest = std::env::var("OXI_LMGRPNEST_DISABLE").is_err();
    // (origin x, origin y, scale x, scale y, rotation deg, centre x, centre y,
    //  flipH, flipV) -- the accumulated turn and mirror of the enclosing groups
    let mut ig_grp: Vec<(f32, f32, f32, f32, f32, f32, f32, bool, bool)> = Vec::new();
    let mut in_grp_pr = false;
    let mut gg = (0.0f32, 0.0f32, 0.0f32, 0.0f32, 0.0f32, 0.0f32, 0.0f32, 0.0f32);
    let mut gg_rot: f32 = 0.0;
    let mut gg_flip = (false, false);
    let mut ig_fill_img = false;
    let mut ig_rot: f32 = 0.0;
    let mut ig_flip = (false, false);
    // A shape whose only unsupported feature is an effect is still emitted --
    // without the effect. Setting this restores the old outright rejection.
    let s_effectshape = std::env::var("OXI_EFFECTSHAPE_DISABLE").is_err();
    let mut ig_paths: Vec<GeomPath> = Vec::new();
    let mut ig_bad = false;
    let mut ig_cur: Option<GeomPath> = None;
    let mut ig_pending: Option<(&'static str, Vec<(f32, f32)>)> = None;
    let mut ig_in = false;
    let mut gr_in = false;
    let mut gr_in_gs = false;
    let mut gr_in_path = false;
    let mut gr_pos: f32 = 0.0;
    let mut gr_color: Option<String> = None;
    let mut gr_alpha: f32 = 1.0;
    let mut gr_stops: Vec<SlideGradientStop> = Vec::new();
    let mut gr_angle: Option<f32> = None;
    let mut gr_scaled = false;
    let mut gr_rot_with_shape = true;
    let mut gr_focus: Option<(f32, f32)> = None;

    let pct = |v: Option<String>| -> f32 {
        v.and_then(|s| s.parse::<f32>().ok())
            .map(|x| x / 100000.0)
            .unwrap_or(0.0)
    };

    loop {
        let ev = reader.read_event_into(&mut buf);
        let (e, empty) = match &ev {
            Ok(Event::Start(e)) => (Some(e), false),
            Ok(Event::Empty(e)) => (Some(e), true),
            Ok(Event::End(e)) => {
                let name = local_name(e.name().as_ref());
                match name.as_str() {
                    "spTree" => in_tree = false,
                    "outerShdw" if in_ig_shdw => {
                        in_ig_shdw = false;
                        if let Some((blur, dist, dir)) = ig_shdw_draft.take() {
                            ig_shadow = Some(crate::ir::ShapeShadow {
                                blur_pt: blur,
                                dist_pt: dist,
                                dir_deg: dir,
                                color: ig_shdw_color
                                    .take()
                                    .unwrap_or_else(|| "000000".to_string()),
                                alpha: ig_shdw_alpha,
                            });
                        }
                    }
                    "grpSpPr" if in_grp_pr => {
                        in_grp_pr = false;
                        let base = if ig_grp.len() >= 2 {
                            ig_grp[ig_grp.len() - 2]
                        } else {
                            (0.0, 0.0, 1.0, 1.0, 0.0, 0.0, 0.0, false, false)
                        };
                        // The group's own box sits in the PARENT's child space,
                        // so it is placed exactly like a shape: translate and
                        // scale, then swing about the parent's centre. Leaving
                        // that swing out fanned d19's top pencil stack across
                        // the page -- those 14 sub-groups carry no rotation of
                        // their own, only their parent's -90.
                        let bw = gg.2 * base.2;
                        let bh = gg.3 * base.3;
                        let (bx, by) = if s_grp_nest {
                            place_in_group(
                                base.0 + gg.0 * base.2,
                                base.1 + gg.1 * base.3,
                                bw,
                                bh,
                                &base,
                            )
                        } else {
                            (base.0 + gg.0 * base.2, base.1 + gg.1 * base.3)
                        };
                        let sx = if gg.6 != 0.0 { bw / gg.6 } else { base.2 };
                        let sy = if gg.7 != 0.0 { bh / gg.7 } else { base.3 };
                        if let Some(top) = ig_grp.last_mut() {
                            *top = (
                                bx - gg.4 * sx,
                                by - gg.5 * sy,
                                sx,
                                sy,
                                // A rotation inherited THROUGH a mirror turns
                                // the other way: for a mirror F and rotation R,
                                // F R(t) = R(-t) F, so the accumulated
                                // transform R_total F_total takes a new group's
                                // own rotation negated whenever the flips so far
                                // are an odd number of mirrors. Adding the two
                                // component-wise instead made d19's pencil
                                // stack read rot -180 -- two groups at rot=-90
                                // flipH=1, which compose to the IDENTITY -- and
                                // drew all 29 pencils upside down, blunt end up.
                                // The leaf path already conjugates a shape's own
                                // rotation this way (`if gfh ^ gfv`); the entry
                                // did not.
                                base.4
                                    + if grpfliprot_on() && (base.7 ^ base.8) {
                                        -gg_rot
                                    } else {
                                        gg_rot
                                    },
                                bx + bw / 2.0,
                                by + bh / 2.0,
                                base.7 ^ (gg_flip.0 && s_grp_nest),
                                base.8 ^ (gg_flip.1 && s_grp_nest),
                            );
                        }
                    }
                    "grpSp" if s_inherit_grp && in_tree && nest == 0 => {
                        ig_grp.pop();
                    }
                    "sp" | "pic" | "grpSp" | "cxnSp" | "graphicFrame" if in_tree => {
                        nest = nest.saturating_sub(1);
                        if nest == 0 {
                            // A shape inside a p:grpSp states its geometry in
                            // the group's child space.
                            if let Some(&entry) = ig_grp.last() {
                                let (ox, oy, sx, sy, grot, _, _, gfh, gfv) = entry;
                                x = ox + x * sx;
                                y = oy + y * sy;
                                w *= sx;
                                h *= sy;
                                let placed = place_in_group(x, y, w, h, &entry);
                                x = placed.0;
                                y = placed.1;
                                ig_flip = (ig_flip.0 ^ gfh, ig_flip.1 ^ gfv);
                                if gfh ^ gfv {
                                    ig_rot = -ig_rot;
                                }
                                if grot.abs() > 1e-4 {
                                    // The group turns as a whole (S-GRPROT):
                                    // the child's centre swings about the
                                    // group's and the child keeps its size,
                                    // turned by the same angle on top of its
                                    // own. d10's layout2 has 12 groups, SIX of
                                    // them rotated -- placing their children
                                    // square left a second, offset copy of the
                                    // frame beside the right one.
                                    ig_rot += grot;
                                }
                            }
                            // Close the candidate.
                            if ok && have_off && have_ext && w > 0.0 && h > 0.0 {
                                let content = match kind {
                                    _ if ig_fill_img && embed.is_some() => {
                                        embed.take().and_then(|rid| {
                                            let rx = archive.try_read_part(rels_path).ok()??;
                                            let rels = parse_relationships(&rx).ok()?;
                                            let rel = rels.get(&rid)?;
                                            let p = resolve_slide_relative_path(
                                                rels_path,
                                                &rel.target,
                                            );
                                            let data = archive.read_binary_part(&p).ok()?;
                                            if data.is_empty() {
                                                return None;
                                            }
                                            Some(ShapeContent::Image {
                                                data,
                                                content_type: detect_content_type(&rel.target),
                                            })
                                        })
                                    }
                                    Some("pic") => embed.take().and_then(|rid| {
                                        let rx = archive.try_read_part(rels_path).ok()??;
                                        let rels = parse_relationships(&rx).ok()?;
                                        let rel = rels.get(&rid)?;
                                        let p = resolve_slide_relative_path(rels_path, &rel.target);
                                        let data = archive.read_binary_part(&p).ok()?;
                                        if data.is_empty() {
                                            return None;
                                        }
                                        // The renderer composites per-pixel
                                        // alpha, so a transparent picture is
                                        // no longer held back. The flag
                                        // restores the old rejection.
                                        if std::env::var("OXI_LMPICALPHA_DISABLE").is_ok()
                                            && media_has_alpha(&data)
                                        {
                                            return None;
                                        }
                                        Some(ShapeContent::Image {
                                            data,
                                            content_type: detect_content_type(&rel.target),
                                        })
                                    }),
                                    // A gradient-only shape has ink of its
                                    // own. Emitting it needed per-stop alpha
                                    // first: d06's layout wash is 020F2B at
                                    // 33.7% over 010C16 at 0%, and emitting it
                                    // while the painter was still opaque laid a
                                    // navy slab over the slide -- 0.8800 ->
                                    // 0.6489.
                                    _ if fill.is_some()
                                        || (s_gradshape && gr_stops.len() >= 2) =>
                                    {
                                        Some(ShapeContent::AutoShape {
                                            paragraphs: Vec::new(),
                                        })
                                    }
                                    _ => None,
                                };
                                if let Some(content) = content {
                                    // A line with <a:noFill/> paints nothing --
                                    // 14 of the 17 accepted rects spell their
                                    // outline that way. (The main shape walker
                                    // sets border_width from a:ln@w regardless,
                                    // a pre-existing bug this must not copy.)
                                    let (bc, bw) = if ln_no_fill || ln_color.is_none() {
                                        (None, None)
                                    } else {
                                        (ln_color.clone(), ln_width)
                                    };
                                    out.push(Shape {
                                        // Inherited (layout/master) shapes are
                                        // drawn but not editable; top-level is
                                        // the harmless default for a non-target.
                                        top_level: true,
                                        // This builder has no view of the
                                        // document-order count, so it declines
                                        // to name one rather than guess.
                                        sp_index: None,
                                        // Group members are walked without a
                                        // bodyPr, and no corpus group asks for
                                        // wrap="none".
                                        wrap_text: true,
                            spc_first_last_para: false,
                                        text_warp: None,
                                        x,
                                        y,
                                        width: w,
                                        height: h,
                                        rotation: ig_rot,
                                        flip_h: ig_flip.0,
                                        flip_v: ig_flip.1,
                                        shape_type: prst.clone(),
                                        ph_type: None,
                                        // Inherited shapes reject prstGeom, so
                                        // there are no adjust values to carry.
                                        adjustments: std::collections::HashMap::new(),
                                        content,
                                        fill_color: if kind == Some("pic") || ig_fill_img {
                                            None
                                        } else {
                                            fill.clone()
                                        },
                                        fill_alpha: if kind == Some("pic") {
                                            None
                                        } else {
                                            fill_alpha
                                        },
                                        border_color: bc,
                                        border_alpha: ln_alpha.take(),
                                        border_width: bw,
                                        // Group members carry no dash yet; the
                                        // corpus states prstDash on top-level
                                        // shapes and connectors only.
                                        border_dash: None,
                                        // A group member's `ln` is walked for
                                        // its colour and width but not for its
                                        // ends, so 35 of the corpus's 359
                                        // decorated ends are dropped here --
                                        // 22 of them on d02 slide 22.
                                        head_end: None,
                                        tail_end: None,
                                        line_cap: None,
                                        l_ins: default_l_ins(),
                                        r_ins: default_r_ins(),
                                        t_ins: default_t_ins(),
                                        b_ins: default_b_ins(),
                                        anchor: None,
                                        src_rect,
                                        fill_rect,
                                        rot_with_shape: true,
                                        image_alpha: ig_img_alpha.take(),
                                        gradient: if gr_stops.len() >= 2 {
                                            Some(SlideGradient {
                                                stops: std::mem::take(&mut gr_stops),
                                                angle_deg: gr_angle,
                                                scaled: gr_scaled,
                                                focus: gr_focus,
                                                rot_with_shape: gr_rot_with_shape,
                                            })
                                        } else {
                                            gr_stops.clear();
                                            None
                                        },
                                        ph_levels: Vec::new(),
                                        // custGeom is accepted now that its
                                        // paths can be drawn (S-CUSTGEOM).
                                        custom_geometry: if ig_bad
                                            || ig_paths.iter().all(|p| p.commands.is_empty())
                                        {
                                            ig_paths.clear();
                                            None
                                        } else {
                                            Some(CustomGeometry {
                                                paths: std::mem::take(&mut ig_paths),
                                                unsupported: false,
                                            })
                                        },
                                        shadow: ig_shadow.take(),
                                        });
                                }
                            }
                            kind = None;
                            ok = false;
                        }
                    }
                    "spPr" if nest > 0 => in_sp_pr = false,
                    "moveTo" | "lnTo" | "cubicBezTo" if ig_in => {
                        if let Some((k, pts)) = ig_pending.take() {
                            match (k, pts.as_slice()) {
                                ("moveTo", [(px, py)]) => {
                                    if let Some(c) = ig_cur.as_mut() {
                                        c.commands.push(GeomCmd::MoveTo(*px, *py));
                                    }
                                }
                                ("lnTo", [(px, py)]) => {
                                    if let Some(c) = ig_cur.as_mut() {
                                        c.commands.push(GeomCmd::LineTo(*px, *py));
                                    }
                                }
                                ("cubicBezTo", [(a1, b1), (a2, b2), (a3, b3)]) => {
                                    if let Some(c) = ig_cur.as_mut() {
                                        c.commands
                                            .push(GeomCmd::CubicTo(*a1, *b1, *a2, *b2, *a3, *b3));
                                    }
                                }
                                _ => ig_bad = true,
                            }
                        }
                    }
                    "path" if ig_in => {
                        if let Some(c) = ig_cur.take() {
                            ig_paths.push(c);
                        }
                    }
                    "custGeom" if ig_in => {
                        ig_in = false;
                        ig_pending = None;
                    }
                    "gs" if gr_in_gs => {
                        gr_in_gs = false;
                        if let Some(c) = gr_color.take() {
                            gr_stops.push(SlideGradientStop {
                                pos: gr_pos,
                                color: c,
                                alpha: gr_alpha,
                            });
                        }
                    }
                    "path" if gr_in_path => gr_in_path = false,
                    "gradFill" if gr_in => gr_in = false,
                    "blipFill" if nest > 0 => in_pic_blip = false,
                    "ln" if nest > 0 => ln_depth = ln_depth.saturating_sub(1),
                    "solidFill" if nest > 0 => in_solid = false,
                    _ => {}
                }
                buf.clear();
                continue;
            }
            Ok(Event::Eof) => break,
            Err(_) => break,
            _ => {
                buf.clear();
                continue;
            }
        };
        let e = match e {
            Some(e) => e,
            None => {
                buf.clear();
                continue;
            }
        };
        let name = local_name(e.name().as_ref());

        if name == "spTree" {
            in_tree = true;
            buf.clear();
            continue;
        }
        if !in_tree {
            buf.clear();
            continue;
        }

        match name.as_str() {
            "grpSp" if s_inherit_grp && nest == 0 => {
                if !empty {
                    let base = ig_grp
                        .last()
                        .copied()
                        .unwrap_or((0.0, 0.0, 1.0, 1.0, 0.0, 0.0, 0.0, false, false));
                    ig_grp.push(base);
                }
            }
            "grpSpPr" if s_inherit_grp && !ig_grp.is_empty() => {
                if !empty {
                    in_grp_pr = true;
                    gg = (0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0);
                    gg_rot = 0.0;
                    gg_flip = (false, false);
                }
            }
            "xfrm" if in_grp_pr => {
                gg_rot = get_attr(e, "rot")
                    .and_then(|v| v.parse::<f32>().ok())
                    .map(|v| v / 60000.0)
                    .unwrap_or(0.0);
                gg_flip = (
                    get_attr(e, "flipH").as_deref() == Some("1"),
                    get_attr(e, "flipV").as_deref() == Some("1"),
                );
            }
            "off" if in_grp_pr => {
                gg.0 = get_attr(e, "x").and_then(|v| v.parse().ok()).map(emu_to_pt).unwrap_or(0.0);
                gg.1 = get_attr(e, "y").and_then(|v| v.parse().ok()).map(emu_to_pt).unwrap_or(0.0);
            }
            "ext" if in_grp_pr => {
                gg.2 = get_attr(e, "cx").and_then(|v| v.parse().ok()).map(emu_to_pt).unwrap_or(0.0);
                gg.3 = get_attr(e, "cy").and_then(|v| v.parse().ok()).map(emu_to_pt).unwrap_or(0.0);
            }
            "chOff" if in_grp_pr => {
                gg.4 = get_attr(e, "x").and_then(|v| v.parse().ok()).map(emu_to_pt).unwrap_or(0.0);
                gg.5 = get_attr(e, "y").and_then(|v| v.parse().ok()).map(emu_to_pt).unwrap_or(0.0);
            }
            "chExt" if in_grp_pr => {
                gg.6 = get_attr(e, "cx").and_then(|v| v.parse().ok()).map(emu_to_pt).unwrap_or(0.0);
                gg.7 = get_attr(e, "cy").and_then(|v| v.parse().ok()).map(emu_to_pt).unwrap_or(0.0);
            }
            "sp" | "pic" | "grpSp" | "cxnSp" | "graphicFrame" => {
                if nest == 0 {
                    // Only p:sp and p:pic can be faithful; a grpSp/cxnSp/
                    // graphicFrame subtree is walked but never emitted.
                    kind = match name.as_str() {
                        "sp" => Some("sp"),
                        "pic" => Some("pic"),
                        _ => None,
                    };
                    ok = kind.is_some();
                    x = 0.0;
                    y = 0.0;
                    w = 0.0;
                    h = 0.0;
                    have_off = false;
                    have_ext = false;
                    prst = None;
                    fill = None;
                    fill_alpha = None;
                    embed = None;
                    src_rect = None;
                    fill_rect = None;
                    ln_color = None;
                    ln_alpha = None;
                    ln_width = None;
                    ln_no_fill = false;
                    // ★Reset the geometry / gradient accumulators per CANDIDATE,
                    // not only when one is consumed: a REJECTED shape is never
                    // pushed, so whatever it accumulated leaked into the next
                    // accepted one. d06 slide 21 showed it as the layout's
                    // top-right gradient painted over the full-slide white
                    // panel, costing that deck 0.0387.
                    ig_paths.clear();
                    ig_bad = false;
                    ig_cur = None;
                    ig_pending = None;
                    ig_in = false;
                    ig_rot = 0.0;
                    ig_flip = (false, false);
                    ig_img_alpha = None;
                    ig_fill_img = false;
                    in_ig_shdw = false;
                    ig_shdw_draft = None;
                    ig_shdw_color = None;
                    ig_shdw_alpha = 1.0;
                    ig_shadow = None;
                    gr_stops.clear();
                    gr_in = false;
                    gr_in_gs = false;
                    gr_in_path = false;
                    gr_color = None;
                    gr_angle = None;
                    gr_scaled = false;
                    gr_focus = None;
                    in_sp_pr = false;
                    in_pic_blip = false;
                    ln_depth = 0;
                    in_solid = false;
                }
                if !empty {
                    nest += 1;
                }
            }
            _ if nest == 0 || kind.is_none() => {}
            // A placeholder is positioned and filled by the placeholder chain,
            // not by this walk.
            "ph" => ok = false,
            "spPr" => {
                if !empty {
                    in_sp_pr = true;
                }
            }
            // p:blipFill (the picture) is a SIBLING of p:spPr; a:blipFill (a
            // shape fill) is inside it, and is a disqualifier below.
            "blipFill" => {
                // `p:blipFill` is a SIBLING of `p:spPr` and makes the shape a
                // picture; `a:blipFill` INSIDE `p:spPr` is a picture used as
                // the shape's fill. The second used to disqualify the shape --
                // but a picture fill is drawn clipped to the shape's own
                // outline (S-BLIPCLIP), so the same `blip`/`srcRect`/`fillRect`
                // handling serves both. 269 layout shapes in the corpus were
                // being dropped for it (d28 59, d03 57, d22 34, d25 31, d21 28).
                if in_sp_pr && !s_inherit_fillimg {
                    ok = false;
                } else if !empty {
                    in_pic_blip = true;
                    ig_fill_img = in_sp_pr;
                }
            }
            "ln" => {
                if let Some(v) = get_attr(e, "w") {
                    ln_width = v.parse::<f32>().ok().map(emu_to_pt);
                }
                if !empty {
                    ln_depth += 1;
                }
            }
            "xfrm" => {
                // A turned or mirrored shape used to be refused outright,
                // because the paint was axis-aligned and would have drawn it
                // straight. It is not any more: `emit_geom_path_gdi`,
                // `draw_preset_shape_gdi` and `transform_picture` all honour
                // `rotation` / `flip_h` / `flip_v`, so the shape can be carried
                // through instead. 64 otherwise-drawable layout shapes across
                // six decks were being dropped for this (d39 16, d40 16,
                // d10 13, d06 7, d24 7, d19 5).
                let r = get_attr(e, "rot")
                    .and_then(|v| v.parse::<f32>().ok())
                    .map(|v| v / 60000.0)
                    .unwrap_or(0.0);
                let fh = get_attr(e, "flipH").as_deref() == Some("1");
                let fv = get_attr(e, "flipV").as_deref() == Some("1");
                if s_inherit_rot {
                    ig_rot = r;
                    ig_flip = (fh, fv);
                } else if r != 0.0 || fh || fv {
                    ok = false;
                }
            }
            "off" if in_sp_pr && !have_off => {
                x = get_attr(e, "x")
                    .and_then(|v| v.parse::<f32>().ok())
                    .map(emu_to_pt)
                    .unwrap_or(0.0);
                y = get_attr(e, "y")
                    .and_then(|v| v.parse::<f32>().ok())
                    .map(emu_to_pt)
                    .unwrap_or(0.0);
                have_off = true;
            }
            "ext" if in_sp_pr && !have_ext => {
                w = get_attr(e, "cx")
                    .and_then(|v| v.parse::<f32>().ok())
                    .map(emu_to_pt)
                    .unwrap_or(0.0);
                h = get_attr(e, "cy")
                    .and_then(|v| v.parse::<f32>().ok())
                    .map(emu_to_pt)
                    .unwrap_or(0.0);
                have_ext = true;
            }
            // A preset the renderer cannot draw would come out as a
            // rectangle, so it is still excluded -- but `draw_preset_shape_gdi`
            // handles four of them, and d16's layouts hold 17 ellipses that
            // were being dropped for being "not rect".
            "prstGeom" => match get_attr(e, "prst") {
                Some(p) if p == "rect" => prst = Some(p),
                Some(p)
                    if s_inherit_prst
                        && matches!(
                            p.as_str(),
                            "ellipse" | "roundRect" | "homePlate" | "teardrop"
                        ) =>
                {
                    prst = Some(p)
                }
                _ => ok = false,
            },
            "custGeom" if s_inherit_geom => {
                ig_in = true;
                ig_paths.clear();
                ig_bad = false;
                ig_cur = None;
                ig_pending = None;
            }
            "custGeom" => ok = false,
            "path" if ig_in => {
                ig_cur = Some(GeomPath {
                    w: get_attr(e, "w").and_then(|v| v.parse::<f32>().ok()).unwrap_or(0.0),
                    h: get_attr(e, "h").and_then(|v| v.parse::<f32>().ok()).unwrap_or(0.0),
                    fill_none: get_attr(e, "fill").as_deref() == Some("none"),
                    commands: Vec::new(),
                });
            }
            "moveTo" | "lnTo" | "cubicBezTo" if ig_in => {
                ig_pending = Some((
                    match name.as_str() {
                        "moveTo" => "moveTo",
                        "lnTo" => "lnTo",
                        _ => "cubicBezTo",
                    },
                    Vec::new(),
                ));
            }
            "arcTo" | "quadBezTo" if ig_in => ig_bad = true,
            "close" if ig_in => {
                if let Some(c) = ig_cur.as_mut() {
                    c.commands.push(GeomCmd::Close);
                }
            }
            "pt" if ig_pending.is_some() => {
                if let Some((_, pts)) = ig_pending.as_mut() {
                    match (
                        get_attr(e, "x").and_then(|v| v.parse::<f32>().ok()),
                        get_attr(e, "y").and_then(|v| v.parse::<f32>().ok()),
                    ) {
                        (Some(px), Some(py)) => pts.push((px, py)),
                        _ => ig_bad = true,
                    }
                }
            }
            "gradFill" if in_sp_pr && ln_depth == 0 && s_inherit_geom => {
                gr_in = true;
                gr_stops.clear();
                gr_angle = None;
                gr_scaled = false;
                gr_focus = None;
                gr_rot_with_shape = get_attr(&e, "rotWithShape").as_deref() != Some("0");
            }
            "gs" if gr_in => {
                gr_in_gs = true;
                gr_pos = get_attr(e, "pos").and_then(|v| gradient_frac(&v)).unwrap_or(0.0);
                gr_color = None;
                gr_alpha = 1.0;
            }
            "alpha" if gr_in_gs => {
                if let Some(v) = get_attr(e, "val") {
                    if let Ok(a) = v.parse::<f32>() {
                        gr_alpha = (a / 100000.0).clamp(0.0, 1.0);
                    }
                }
            }
            "lin" if gr_in => {
                gr_angle = get_attr(e, "ang")
                    .and_then(|v| v.parse::<f32>().ok())
                    .map(|v| v / 60_000.0);
                gr_scaled = get_attr(e, "scaled").as_deref() == Some("1");
            }
            "path" if gr_in => {
                if get_attr(e, "path").as_deref() == Some("circle") {
                    gr_in_path = true;
                    gr_focus = Some((0.5, 0.5));
                }
            }
            "fillToRect" if gr_in_path => {
                let l = get_attr(e, "l").and_then(|v| gradient_frac(&v)).unwrap_or(0.5);
                let t = get_attr(e, "t").and_then(|v| gradient_frac(&v)).unwrap_or(0.5);
                let r = get_attr(e, "r").and_then(|v| gradient_frac(&v)).unwrap_or(0.5);
                let b = get_attr(e, "b").and_then(|v| gradient_frac(&v)).unwrap_or(0.5);
                gr_focus = Some(((l + (1.0 - r)) / 2.0, (t + (1.0 - b)) / 2.0));
            }
            "gradFill" | "pattFill" if in_sp_pr && ln_depth == 0 => ok = false,
            "solidFill" if in_sp_pr && ln_depth == 0 => {
                if !empty {
                    in_solid = true;
                }
            }
            "solidFill" if ln_depth > 0 => {
                if !empty {
                    in_solid = true;
                }
            }
            "noFill" if ln_depth > 0 => ln_no_fill = true,
            // A constant fill alpha IS reproducible now (S-FILLALPHA), so it is
            // captured rather than rejected -- but only on the shape's OWN
            // solidFill, and only while the renderer will composite it. An
            // alpha anywhere else (a text colour, say) still has no
            // representation, so the shape is still dropped, and so is one
            // whose @val does not parse.
            "alpha" if in_solid && ln_depth == 0 => {
                match get_attr(e, "val")
                    .and_then(|v| v.parse::<f32>().ok())
                    .filter(|_| {
                        in_sp_pr && std::env::var("OXI_FILLALPHA_DISABLE").is_err()
                    }) {
                    Some(p) => fill_alpha = Some((p / 100000.0).clamp(0.0, 1.0)),
                    None => ok = false,
                }
            }
            "lumMod" | "lumOff" | "tint" | "shade" | "satMod" | "hueMod"
                if in_solid && ln_depth == 0 =>
            {
                ok = false;
            }
            "srgbClr" | "schemeClr" if in_ig_shdw => {
                let val = get_attr(e, "val").unwrap_or_default();
                ig_shdw_color = Some(if name == "srgbClr" {
                    val
                } else {
                    theme_colors
                        .get(&val)
                        .cloned()
                        .unwrap_or_else(|| scheme_color_to_hex(&val))
                });
            }
            "alpha" if ln_depth > 0 => {
                if let Some(v) = get_attr(e, "val").and_then(|v| v.parse::<f32>().ok()) {
                    ln_alpha = Some((v / 100000.0).clamp(0.0, 1.0));
                }
            }
            "alpha" if in_ig_shdw => {
                if let Some(v) = get_attr(e, "val").and_then(|v| v.parse::<f32>().ok()) {
                    ig_shdw_alpha = (v / 100000.0).clamp(0.0, 1.0);
                }
            }
            "srgbClr" | "schemeClr" if gr_in_gs && gr_color.is_none() => {
                gr_color = if name == "srgbClr" {
                    get_attr(e, "val")
                } else {
                    get_attr(e, "val").map(|v| {
                        theme_colors
                            .get(&v)
                            .cloned()
                            .unwrap_or_else(|| scheme_color_to_hex(&v))
                    })
                };
            }
            "srgbClr" | "schemeClr" if in_solid => {
                let hex = if name == "srgbClr" {
                    get_attr(e, "val")
                } else {
                    get_attr(e, "val").map(|v| {
                        theme_colors
                            .get(&v)
                            .cloned()
                            .unwrap_or_else(|| scheme_color_to_hex(&v))
                    })
                };
                if ln_depth > 0 {
                    if ln_color.is_none() {
                        ln_color = hex;
                    }
                } else if in_sp_pr && fill.is_none() {
                    fill = hex;
                }
            }
            "blip" if in_pic_blip => {
                for attr in e.attributes().flatten() {
                    let key = std::str::from_utf8(attr.key.as_ref()).unwrap_or("");
                    if key == "r:embed" || key.ends_with(":embed") {
                        embed = Some(String::from_utf8_lossy(&attr.value).to_string());
                    }
                }
            }
            // `@amt` is the picture's opacity. The blit CAN reproduce it now
            // (`image_alpha` / `alpha_blit`), so carry it instead of refusing
            // the shape -- 46 layout pictures in the corpus were being dropped
            // for it (d04 34, d06 12).
            "alphaModFix" if in_pic_blip => {
                match get_attr(e, "amt").and_then(|a| a.parse::<f32>().ok()) {
                    Some(a) if s_inherit_imgalpha => {
                        ig_img_alpha = Some((a / 100000.0).clamp(0.0, 1.0));
                    }
                    Some(a) if a != 100000.0 => ok = false,
                    _ => {}
                }
            }
            "duotone" | "clrChange" | "grayscl" | "biLevel" | "tile" if in_pic_blip => ok = false,
            "srcRect" if in_pic_blip => {
                src_rect = Some((
                    pct(get_attr(e, "l")),
                    pct(get_attr(e, "t")),
                    pct(get_attr(e, "r")),
                    pct(get_attr(e, "b")),
                ));
            }
            "fillRect" if in_pic_blip => {
                fill_rect = Some((
                    pct(get_attr(e, "l")),
                    pct(get_attr(e, "t")),
                    pct(get_attr(e, "r")),
                    pct(get_attr(e, "b")),
                ));
            }
            // A shadow/glow changes what the shape looks like well beyond its
            // box, so the shape used not to be emitted at all. That trades a
            // missing soft edge for a missing shape, which is the larger error
            // when the shape is big: d24's layout draws three full-height
            // gradient bands, each carrying an `a:outerShdw`, and Oxi drew bare
            // background where PowerPoint draws half the slide.
            "outerShdw" | "innerShdw" | "softEdge" | "reflection" | "glow"
                if !s_effectshape =>
            {
                ok = false
            }
            "outerShdw" if in_sp_pr => {
                let emu = |name: &str| {
                    get_attr(e, name)
                        .and_then(|v| v.parse::<f32>().ok())
                        .unwrap_or(0.0)
                };
                let draft = (emu("blurRad") / 12700.0, emu("dist") / 12700.0, emu("dir") / 60000.0);
                if empty {
                    // Self-closing: nothing but attributes; black is the
                    // schema's default shadow colour.
                    ig_shadow = Some(crate::ir::ShapeShadow {
                        blur_pt: draft.0,
                        dist_pt: draft.1,
                        dir_deg: draft.2,
                        color: "000000".to_string(),
                        alpha: 1.0,
                    });
                } else {
                    in_ig_shdw = true;
                    ig_shdw_draft = Some(draft);
                    ig_shdw_color = None;
                    ig_shdw_alpha = 1.0;
                }
            }
            _ => {}
        }
        buf.clear();
    }

    out
}

/// `p:cSld/p:bg/p:bgPr/a:blipFill/a:blip@r:embed` -> the relationship id of the
/// background picture, or None when the background is not a picture fill.
///
/// Only the id is returned: resolving it needs the rels of the part the `p:bg`
/// was found in (slide, layout or master), which differ, so the caller does the
/// lookup. `a:blip` also appears inside shape fills, hence the `in_bg_pr` gate;
/// the walk stops at `</p:bg>` for the same reason.
fn parse_bg_blip_rid(xml: &str) -> Option<String> {
    let mut reader = Reader::from_str(xml);
    let mut in_bg = false;
    let mut in_bg_pr = false;
    let mut found: Option<String> = None;
    loop {
        let ev = match reader.read_event() {
            Ok(e) => e,
            Err(_) => break,
        };
        match ev {
            Event::Start(ref e) | Event::Empty(ref e) => {
                match local_name(e.name().as_ref()).as_str() {
                    "bg" => in_bg = true,
                    "bgPr" if in_bg => in_bg_pr = true,
                    "blip" if in_bg_pr && found.is_none() => {
                        for attr in e.attributes().flatten() {
                            let key = std::str::from_utf8(attr.key.as_ref()).unwrap_or("");
                            if key == "r:embed" || key.ends_with(":embed") {
                                found = Some(String::from_utf8_lossy(&attr.value).to_string());
                            }
                        }
                    }
                    _ => {}
                }
            }
            Event::End(e) => match local_name(e.name().as_ref()).as_str() {
                "bgPr" => in_bg_pr = false,
                "bg" => break,
                _ => {}
            },
            Event::Eof => break,
            _ => {}
        }
    }
    found
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
/// The layout placeholder text styles that apply to a slide shape, using the
/// same key normalisation as `lookup_ph_anchor` (a slide's "ctrTitle" and a
/// layout's "title" are the same slot).
/// Layout placeholder styles laid over master placeholder styles, field by
/// field and level by level. The layout is the nearer ancestor, so anything it
/// states wins; anything it leaves out falls through to the master's
/// placeholder, which in turn beats the master's `p:txStyles`.
/// The master placeholder's face survives an empty layout level unless this
/// is set.
fn phfont_merge_on() -> bool {
    std::env::var("OXI_PHFONTMERGE_DISABLE").is_err()
}

/// A group's rotation is conjugated by the flips above it unless this is set.
fn grpfliprot_on() -> bool {
    std::env::var("OXI_GRPFLIPROT_DISABLE").is_err()
}

fn merge_ph_levels(
    layout: Vec<MasterStyleLevel>,
    master: Vec<MasterStyleLevel>,
) -> Vec<MasterStyleLevel> {
    if layout.is_empty() {
        return master;
    }
    if master.is_empty() {
        return layout;
    }
    let n = layout.len().max(master.len());
    (0..n)
        .map(|i| {
            let l = layout.get(i.min(layout.len() - 1)).cloned().unwrap_or_default();
            let m = master.get(i.min(master.len() - 1)).cloned().unwrap_or_default();
            MasterStyleLevel {
                font_size: l.font_size.or(m.font_size),
                color: l.color.clone().or(m.color.clone()),
                algn: l.algn.or(m.algn),
                line_spacing: l.line_spacing.or(m.line_spacing),
                // The face has to fall through too, and `..l` alone dropped
                // it: d24's LAYOUT title placeholder declares an lstStyle with
                // nothing but an empty `<a:defRPr/>`, and that empty level was
                // enough to discard the MASTER placeholder's
                // `<a:latin typeface="Fira Sans SemiBold"/>`, leaving the
                // title on the theme's Arial.
                font_family: if phfont_merge_on() {
                    l.font_family.clone().or(m.font_family.clone())
                } else {
                    l.font_family.clone()
                },
                highlight: l.highlight.clone().or(m.highlight.clone()),
                italic: l.italic || m.italic,
                bold: l.bold.or(m.bold),
                ..l
            }
        })
        .collect()
}

/// The placeholder levels a shape inherits from one map.
///
/// `any_idx` is for the MASTER map only. A master carries ONE placeholder per
/// type and every slide placeholder of that type inherits it whatever its idx
/// says, which is what d24 slide 22 needs: its shape is
/// `<p:ph idx="4294967295" type="body"/>` -- the sentinel PowerPoint writes for
/// an unset idx -- layout10 has no body placeholder at all, and the master's
/// is `idx="1"`. PowerPoint drew that paragraph at exactly 24.00pt, the master
/// PLACEHOLDER's `sz="2400"`, not the 14pt its `p:txStyles/p:bodyStyle` says.
///
/// A LAYOUT may hold several placeholders of one type with different styles,
/// so matching by type alone there would pick an arbitrary one.
fn lookup_ph_levels_in(
    layout: &HashMap<(Option<String>, Option<String>), Vec<MasterStyleLevel>>,
    ph_type: Option<&String>,
    ph_idx: Option<&String>,
    any_idx: bool,
) -> Vec<MasterStyleLevel> {
    let mut keys: Vec<(Option<String>, Option<String>)> = Vec::new();
    keys.push((ph_type.cloned(), ph_idx.cloned()));
    if let Some(ty) = ph_type {
        if ty == "ctrTitle" {
            keys.push((Some("title".to_string()), ph_idx.cloned()));
        }
        if ty == "title" {
            keys.push((Some("ctrTitle".to_string()), ph_idx.cloned()));
        }
        // S-SUBTITLE (2026-08-25): a `subTitle` inherits from the master's
        // `body` placeholder, the same way `ctrTitle` inherits from `title`.
        // Only the title half of that pair was here, so a subTitle whose own
        // layout does not declare one fell all the way through to the THEME
        // face. d16 slide 25's "Any questions?" is the specimen: PowerPoint
        // sets it in Source Sans Pro Bold (the master body placeholder's face,
        // 496px wide) and Oxi set it in Arial.
        //
        // **50 subTitle placeholders carrying text over 8 dev decks** reach
        // nothing in their own layout or master by exact type (against 51
        // ctrTitle, which this alias table already rescued).
        if ty == "subTitle" && subtitle_alias_on() {
            keys.push((Some("body".to_string()), ph_idx.cloned()));
        }
        keys.push((Some(ty.clone()), None));
    }
    if let Some(idx) = ph_idx {
        let idx = idx.clone();
        keys.push((Some("body".to_string()), Some(idx.clone())));
        keys.push((Some("obj".to_string()), Some(idx.clone())));
        keys.push((None, Some(idx)));
    }
    for k in &keys {
        if let Some(v) = layout.get(k) {
            return v.clone();
        }
    }
    if any_idx && phanyidx_on() {
        if let Some(ty) = ph_type {
            // `title` and `ctrTitle` name the same slot, and the exact-key list
            // above only pairs the alias with the shape's OWN idx. d35's slide
            // shape is `ctrTitle idx=4294967295` while its master declares one
            // `<p:ph type="title"/>` with no idx, so neither the alias key nor a
            // same-type sweep reaches it -- and the level it carries is the
            // white `a:highlight` behind BIG CONCEPT and the condensed Oswald
            // that lets PowerPoint fit the line in 473.7pt of a 553pt box. 74
            // title/ctrTitle placeholders over 8 dev decks are reachable only
            // this way.
            let alias = match ty.as_str() {
                _ if !phalias_on() => None,
                "ctrTitle" => Some("title"),
                "title" => Some("ctrTitle"),
                "subTitle" if subtitle_alias_on() => Some("body"),
                _ => None,
            };
            // Deterministic: several entries of one type would otherwise
            // depend on HashMap order. Own type first, then the alias.
            for want in [Some(ty.as_str()), alias].into_iter().flatten() {
                let mut hit: Option<(&Option<String>, &Vec<MasterStyleLevel>)> = None;
                for ((t, i), v) in layout {
                    if t.as_deref() == Some(want) && hit.is_none_or(|(bi, _)| i < bi) {
                        hit = Some((i, v));
                    }
                }
                if let Some((_, v)) = hit {
                    return v.clone();
                }
            }
        }
    }
    Vec::new()
}

/// `title` and `ctrTitle` are matched as the same slot unless this is set.
fn phalias_on() -> bool {
    std::env::var("OXI_PHALIAS_DISABLE").is_err()
}

/// A `subTitle` placeholder inherits the master's `body` levels unless this is
/// set (which restores letting it fall through to the theme face).
fn subtitle_alias_on() -> bool {
    std::env::var("OXI_SUBTITLE_ALIAS_DISABLE").is_err()
}

/// A master placeholder is inherited by type regardless of idx unless this is
/// set.
fn phanyidx_on() -> bool {
    std::env::var("OXI_PHANYIDX_DISABLE").is_err()
}

fn lookup_ph_levels(
    layout: &HashMap<(Option<String>, Option<String>), Vec<MasterStyleLevel>>,
    ph_type: Option<&String>,
    ph_idx: Option<&String>,
) -> Vec<MasterStyleLevel> {
    lookup_ph_levels_in(layout, ph_type, ph_idx, false)
}

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
    let mut hole_size: Option<f64> = None;
    let mut bubble_scale: Option<f64> = None;
    let mut size_represents: Option<String> = None;
    let mut hi_low_lines = false;
    let mut up_down_bars = false;
    let mut up_down_gap: Option<f64> = None;
    let mut in_up_down = false;
    let mut legend_overlay = true;
    let mut in_legend = false;
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
    let mut in_val_ax = false;
    let mut val_min: Option<f64> = None;
    let mut val_max: Option<f64> = None;
    let mut in_dpt = false;
    let mut in_ser_sppr = false;
    let mut dpt_idx: Option<u32> = None;
    let mut dpt_color: Option<String> = None;
    let mut ser_color: Option<String> = None;
    let mut ser_point_colors: Vec<(u32, String)> = Vec::new();
    // Which cache we're collecting: "tx" | "cat" | "val" | ""
    let mut ser_target = "";
    let mut ser_name: Option<String> = None;
    let mut ser_values: Vec<f64> = Vec::new();
    let mut ser_x_values: Vec<f64> = Vec::new();
    let mut ser_sizes: Vec<f64> = Vec::new();
    // Scatter per-series draw flags. Word's render-truth discriminators
    // (chart_scatter probe): <c:spPr><a:ln><a:noFill/> = no connecting line
    // (markers only), <c:marker><c:symbol val="none"/> = no markers.
    let mut ser_line_none = false;
    let mut ser_marker_none = false;
    let mut in_ser_ln = false;
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
                    // A doughnut is a pie with a hole: the same clockwise
                    // slice geometry, one ring per series (Word draws the
                    // first series only, like the pie).
                    "doughnutChart" => {
                        chart_type = Some("doughnut".to_string());
                    }
                    "barChart" => {
                        in_bar_chart = true;
                        chart_type = Some("bar".to_string());
                    }
                    // <c:areaChart> — the grouping child (standard / stacked /
                    // percentStacked) drives the same three regimes the bar
                    // chart uses; the series carry the same strCache/numCache.
                    "areaChart" => {
                        chart_type = Some("area".to_string());
                    }
                    // <c:scatterChart> — the XY chart. Its series carry
                    // c:xVal/c:yVal (numeric on BOTH axes) instead of
                    // c:cat/c:val, so there is no category band at all.
                    "scatterChart" => {
                        chart_type = Some("scatter".to_string());
                    }
                    // <c:bubbleChart> — an XY chart whose points carry a third
                    // number (c:bubbleSize) drawn as the disc radius.
                    "bubbleChart" => {
                        chart_type = Some("bubble".to_string());
                    }
                    // <c:radarChart> — the categories become SPOKES of a
                    // polar web instead of a horizontal band; the sub-style
                    // rides in <c:radarStyle val="marker|filled|standard"/>
                    // (a self-closing child, captured in the Empty arm and
                    // parked in `grouping`, which no radar path reads).
                    "radarChart" => {
                        chart_type = Some("radar".to_string());
                    }
                    // <c:stockChart> — high/low/close (or open/high/low/
                    // close) plotted on the LINE chart's geometry. The
                    // series carry <a:ln><a:noFill/> so nothing connects
                    // the points; the data is carried by <c:hiLowLines>
                    // and <c:upDownBars>, both direct children here.
                    "stockChart" => {
                        chart_type = Some("stock".to_string());
                    }
                    // <c:upDownBars> is a Start element (its gapWidth is
                    // a child); hiLowLines is self-closing (Empty arm).
                    "upDownBars" => {
                        up_down_bars = true;
                        in_up_down = true;
                    }
                    "ser"
                        if in_bar_chart
                            || chart_type.as_deref() == Some("pie")
                            || chart_type.as_deref() == Some("line")
                            || chart_type.as_deref() == Some("doughnut")
                            || chart_type.as_deref() == Some("area")
                            || chart_type.as_deref() == Some("scatter")
                            || chart_type.as_deref() == Some("bubble")
                            || chart_type.as_deref() == Some("radar")
                            || chart_type.as_deref() == Some("stock") =>
                    {
                        in_ser = true;
                        ser_target = "";
                        ser_name = None;
                        ser_values.clear();
                        ser_x_values.clear();
                        ser_sizes.clear();
                        ser_categories.clear();
                        ser_line_none = false;
                        ser_marker_none = false;
                        in_ser_ln = false;
                        ser_color = None;
                        ser_point_colors.clear();
                        in_ser_sppr = false;
                        in_dpt = false;
                    }
                    "valAx" => in_val_ax = true,
                    "dPt" if in_ser => {
                        in_dpt = true;
                        dpt_idx = None;
                        dpt_color = None;
                    }
                    "spPr" if in_ser && !in_dpt && !in_dlbls => in_ser_sppr = true,
                    "ln" if in_ser && !in_dlbls => in_ser_ln = true,
                    "noFill" if in_ser_ln => ser_line_none = true,
                    "tx" if in_ser => ser_target = "tx",
                    "cat" if in_ser => ser_target = "cat",
                    "val" if in_ser => ser_target = "val",
                    "xVal" if in_ser => ser_target = "xval",
                    "yVal" if in_ser => ser_target = "yval",
                    "bubbleSize" if in_ser => ser_target = "bubsize",
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
                    "legend" => {
                        has_legend = true;
                        in_legend = true;
                    }
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
                    // Scatter per-series draw flags. Both discriminators are
                    // SELF-CLOSING (<a:noFill/>, <c:symbol val="none"/>) so
                    // they arrive as Event::Empty, never Start.
                    "noFill" if in_ser_ln => ser_line_none = true,
                    "symbol" if in_ser && !in_dlbls => {
                        if get_attr(&e, "val").as_deref() == Some("none") {
                            ser_marker_none = true;
                        }
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
                    // <c:holeSize val="50"/> — self-closing child of
                    // <c:doughnutChart>, the hole diameter as a percent of
                    // the outer diameter (same Event::Empty trap as barDir).
                    "holeSize" => {
                        if let Some(v) = get_attr(&e, "val") {
                            if let Ok(n) = v.parse::<f64>() {
                                hole_size = Some(n);
                            }
                        }
                    }
                    // <c:min val="0.7"/> / <c:max val="0.9"/> inside
                    // <c:valAx><c:scaling>. Only the VALUE axis is read -- a
                    // category axis carries the same element names -- and the
                    // tag has to match exactly, since `minorTickMark` and
                    // friends share the prefix. Self-closing, so this is the
                    // `Event::Empty` arm and not the `Start` one.
                    "min" if in_val_ax => {
                        if let Some(v) = get_attr(&e, "val") {
                            if let Ok(n) = v.parse::<f64>() {
                                val_min = Some(n);
                            }
                        }
                    }
                    "max" if in_val_ax => {
                        if let Some(v) = get_attr(&e, "val") {
                            if let Ok(n) = v.parse::<f64>() {
                                val_max = Some(n);
                            }
                        }
                    }
                    // <c:idx val="0"/> inside <c:dPt>, and the colour that
                    // follows it. `idx` also names the SERIES, so it is only
                    // read while a data point is open.
                    "idx" if in_dpt => {
                        if let Some(v) = get_attr(&e, "val") {
                            if let Ok(n) = v.parse::<u32>() {
                                dpt_idx = Some(n);
                            }
                        }
                    }
                    "srgbClr" if in_dpt => {
                        if dpt_color.is_none() {
                            dpt_color = get_attr(&e, "val");
                        }
                    }
                    "srgbClr" if in_ser_sppr => {
                        if ser_color.is_none() {
                            ser_color = get_attr(&e, "val");
                        }
                    }
                    // <c:bubbleScale val="200"/> and <c:sizeRepresents
                    // val="w"/> — self-closing children of <c:bubbleChart>
                    // (the same Event::Empty trap as holeSize / barDir).
                    "bubbleScale" => {
                        if let Some(v) = get_attr(&e, "val") {
                            if let Ok(n) = v.parse::<f64>() {
                                bubble_scale = Some(n);
                            }
                        }
                    }
                    "sizeRepresents" => {
                        if let Some(v) = get_attr(&e, "val") {
                            size_represents = Some(v);
                        }
                    }
                    // <c:radarStyle val="filled"/> — self-closing child of
                    // <c:radarChart>. "filled" fills each series polygon;
                    // "marker"/"standard" stroke it only (Word draws NO
                    // markers in any measured arm, so the two are identical).
                    // <c:hiLowLines/> — self-closing; its presence is the
                    // whole switch (Word draws the rule with the default
                    // black w=0.75 when there is no <c:spPr>).
                    "hiLowLines" => {
                        hi_low_lines = true;
                    }
                    // <c:gapWidth> inside <c:upDownBars> (the bar chart's
                    // own gapWidth is handled by its own path).
                    "gapWidth" if in_up_down => {
                        if let Some(v) = get_attr(&e, "val") {
                            if let Ok(n) = v.parse::<f64>() {
                                up_down_gap = Some(n);
                            }
                        }
                    }
                    "radarStyle" => {
                        if let Some(v) = get_attr(&e, "val") {
                            grouping = Some(v);
                        }
                    }
                    // python-pptx writes a bare self-closing <c:legend/> to
                    // enable a legend (no overlay/position attrs). Any legend
                    // declaration -> has_legend.
                    "legend" => has_legend = true,
                    // <c:overlay val="0"/> inside <c:legend>: the
                    // legend is NOT overlaid, so it takes a band on
                    // the right and the plot area shrinks. A bare
                    // <c:legend/> (no overlay child) IS an overlay
                    // and leaves the plot alone — chart_pie2 p2/p3
                    // (bare) keep the circle on the frame centre and
                    // the legend overlaps it, while chart_doughnut
                    // (overlay=0) shifts the ring left. <c:overlay>
                    // also occurs inside <c:title>, hence in_legend.
                    "overlay" if in_legend => {
                        if get_attr(&e, "val").as_deref() == Some("0") {
                            legend_overlay = false;
                        }
                    }
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
                    "legend" => in_legend = false,
                    "upDownBars" => in_up_down = false,
                    "ln" => in_ser_ln = false,
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
                                // A scatter series' y values live in c:yVal;
                                // c:val never appears, so both feed the same
                                // `values` vec and stay mutually exclusive.
                                "val" | "yval" => {
                                    if let Ok(v) = cur_v.trim().parse::<f64>() {
                                        ser_values.push(v);
                                    }
                                }
                                "xval" => {
                                    if let Ok(v) = cur_v.trim().parse::<f64>() {
                                        ser_x_values.push(v);
                                    }
                                }
                                "bubsize" => {
                                    if let Ok(v) = cur_v.trim().parse::<f64>() {
                                        ser_sizes.push(v);
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
                                x_values: std::mem::take(&mut ser_x_values),
                                line_none: ser_line_none,
                                marker_none: ser_marker_none,
                                sizes: std::mem::take(&mut ser_sizes),
                                color: ser_color.take(),
                                point_colors: std::mem::take(&mut ser_point_colors),
                            });
                            in_ser_ln = false;
                        }
                    }
                    "valAx" => in_val_ax = false,
                    "dPt" if in_dpt => {
                        in_dpt = false;
                        if let (Some(i), Some(c)) = (dpt_idx.take(), dpt_color.take()) {
                            ser_point_colors.push((i, c));
                        }
                    }
                    "spPr" if in_ser_sppr => in_ser_sppr = false,
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
        val_min,
        val_max,
        chart_type: chart_type.unwrap_or_else(default_chart_type),
        bar_dir: bar_dir.unwrap_or_else(default_chart_bar_dir),
        grouping: grouping.unwrap_or_else(default_chart_grouping),
        hole_size: hole_size.unwrap_or_else(default_chart_hole_size),
        bubble_scale: bubble_scale.unwrap_or_else(default_chart_bubble_scale),
        size_represents: size_represents
            .unwrap_or_else(default_chart_size_represents),
        hi_low_lines,
        up_down_bars,
        up_down_gap: up_down_gap.unwrap_or_else(default_chart_updown_gap),
        legend_overlay,
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
    let first_slide_num = parse_first_slide_num(&pres_xml);

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

    // 2.9. Fonts the deck carries inside itself. Read before the slides so a
    // consumer can install them before it measures any text.
    let embedded_fonts = if std::env::var("OXI_EMBEDFONT_DISABLE").is_err() {
        parse_embedded_fonts(&pres_xml, &rid_to_path, &mut archive)
    } else {
        Vec::new()
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
                    SlidePos { index: i + 1, first_slide_num },
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
        embedded_fonts,
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

    /// A one-entry archive, enough for `parse_slide` to run: every part it
    /// cannot find (rels, media) is optional.
    fn archive_of(name: &str, body: &str) -> OoxmlArchive {
        use std::io::{Cursor, Write};
        let mut out = Vec::new();
        {
            let mut w = zip::ZipWriter::new(Cursor::new(&mut out));
            w.start_file(name, zip::write::SimpleFileOptions::default())
                .unwrap();
            w.write_all(body.as_bytes()).unwrap();
            w.finish().unwrap();
        }
        OoxmlArchive::new(&out).unwrap()
    }

    fn slide_with(ln_body: &str) -> Slide {
        let xml = format!(
            r#"<?xml version="1.0" encoding="UTF-8"?>
<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
       xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
 <p:cSld><p:spTree>
  <p:cxnSp>
   <p:nvCxnSpPr><p:cNvPr id="2" name="c"/><p:cNvCxnSpPr/><p:nvPr/></p:nvCxnSpPr>
   <p:spPr>
    <a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="0"/></a:xfrm>
    <a:prstGeom prst="straightConnector1"><a:avLst/></a:prstGeom>
    <a:ln w="19050">{ln_body}</a:ln>
   </p:spPr>
  </p:cxnSp>
 </p:spTree></p:cSld>
</p:sld>"#
        );
        let mut ar = archive_of("ppt/slides/slide1.xml", &xml);
        parse_slide(
            &xml,
            SlidePos { index: 0, first_slide_num: 1 },
            &mut ar,
            "ppt/slides/_rels/slide1.xml.rels",
            &HashMap::new(),
            &HashMap::new(),
            &HashMap::new(),
        )
        .expect("slide must parse")
    }

    #[test]
    fn line_ends_are_read_in_both_element_forms() {
        // `a:headEnd` / `a:tailEnd` carry nothing but attributes, so quick-xml
        // hands them to the EMPTY arm; a handler written only on the Start arm
        // is silently inert. Both forms are stated here so neither arm can be
        // dropped without a test going red.
        let s = slide_with(
            r#"<a:headEnd type="oval" w="lg" len="sm"/>
               <a:tailEnd type="triangle"></a:tailEnd>"#,
        );
        let sh = &s.shapes[0];
        let h = sh.head_end.as_ref().expect("headEnd (empty form)");
        assert_eq!((h.kind.as_str(), h.w.as_str(), h.len.as_str()), ("oval", "lg", "sm"));
        let t = sh.tail_end.as_ref().expect("tailEnd (start/end form)");
        // Absent @w / @len mean "med", not "missing".
        assert_eq!((t.kind.as_str(), t.w.as_str(), t.len.as_str()), ("triangle", "med", "med"));
    }

    /// A slide holding one rectangle whose `p:spPr` body is `sp_pr`.
    fn rect_shape_with(xfrm_attrs: &str, sp_pr: &str) -> Slide {
        let xml = format!(
            r#"<?xml version="1.0" encoding="UTF-8"?>
<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
       xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
 <p:cSld><p:spTree>
  <p:sp>
   <p:nvSpPr><p:cNvPr id="2" name="r"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
   <p:spPr>
    <a:xfrm {xfrm_attrs}><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm>
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
    {sp_pr}
   </p:spPr>
  </p:sp>
 </p:spTree></p:cSld>
</p:sld>"#
        );
        let mut ar = archive_of("ppt/slides/slide1.xml", &xml);
        parse_slide(
            &xml,
            SlidePos { index: 0, first_slide_num: 1 },
            &mut ar,
            "ppt/slides/_rels/slide1.xml.rels",
            &HashMap::new(),
            &HashMap::new(),
            &HashMap::new(),
        )
        .expect("slide must parse")
    }

    #[test]
    fn gradient_stop_colours_are_read_in_both_element_forms() {
        // A stop's colour is self-closing whenever it carries no modifier, so
        // it reaches quick-xml on the EMPTY arm; with only the Start-arm
        // handler the stop was dropped, and a ramp short of two stops is
        // discarded whole. 138 shape gradients over 5 dev decks are written
        // this way. Both spellings are stated here so neither arm can go.
        let s = rect_shape_with(
            "",
            r#"<a:gradFill>
                 <a:gsLst>
                   <a:gs pos="0"><a:srgbClr val="000000"/></a:gs>
                   <a:gs pos="100000"><a:srgbClr val="FFFFFF"><a:alpha val="50000"/></a:srgbClr></a:gs>
                 </a:gsLst>
                 <a:lin ang="5400000" scaled="0"/>
               </a:gradFill>"#,
        );
        let g = s.shapes[0].gradient.as_ref().expect("gradient with a bare stop colour");
        assert_eq!(g.stops.len(), 2);
        assert_eq!(g.stops[0].color, "000000");
        assert!((g.stops[0].alpha - 1.0).abs() < 1e-6);
        assert_eq!(g.stops[1].color, "FFFFFF");
        assert!((g.stops[1].alpha - 0.5).abs() < 1e-6);
        // `a:lin` is itself always self-closing -- the same trap, one element up.
        assert!((g.angle_deg.expect("a:lin@ang") - 90.0).abs() < 1e-3);
        // Absent `rotWithShape` means the ramp turns with the shape (gradrot
        // probe block B: absent behaved as "1" on 6 of 6 arms).
        assert!(g.rot_with_shape);
    }

    #[test]
    fn gradient_rot_with_shape_zero_pins_the_ramp() {
        let s = rect_shape_with(
            r#"rot="10800000""#,
            r#"<a:gradFill rotWithShape="0">
                 <a:gsLst>
                   <a:gs pos="0"><a:srgbClr val="112233"/></a:gs>
                   <a:gs pos="100000"><a:srgbClr val="445566"/></a:gs>
                 </a:gsLst>
                 <a:lin ang="0" scaled="0"/>
               </a:gradFill>"#,
        );
        let sh = &s.shapes[0];
        assert!((sh.rotation - 180.0).abs() < 1e-3);
        assert!(!sh.gradient.as_ref().expect("gradient").rot_with_shape);
    }

    #[test]
    fn line_end_type_none_is_no_decoration() {
        // Every table-cell border in the corpus states type="none"; storing
        // those would mean drawing nothing 1410 times.
        let s = slide_with(r#"<a:headEnd type="none" w="sm" len="sm"/><a:tailEnd/>"#);
        assert!(s.shapes[0].head_end.is_none());
        assert!(s.shapes[0].tail_end.is_none());
    }

    /// A slide holding one text box whose single paragraph is `body`.
    fn text_slide_with(body: &str, index: usize, first: u32) -> Slide {
        let xml = format!(
            r#"<?xml version="1.0" encoding="UTF-8"?>
<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
       xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
 <p:cSld><p:spTree>
  <p:sp>
   <p:nvSpPr><p:cNvPr id="2" name="t"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>
   <p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm>
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>
   <p:txBody><a:bodyPr/><a:lstStyle/><a:p>{body}</a:p></p:txBody>
  </p:sp>
 </p:spTree></p:cSld>
</p:sld>"#
        );
        let mut ar = archive_of("ppt/slides/slide1.xml", &xml);
        parse_slide(
            &xml,
            SlidePos { index, first_slide_num: first },
            &mut ar,
            "ppt/slides/_rels/slide1.xml.rels",
            &HashMap::new(),
            &HashMap::new(),
            &HashMap::new(),
        )
        .expect("slide must parse")
    }

    fn only_text(s: &Slide) -> String {
        match &s.shapes[0].content {
            ShapeContent::TextBox { paragraphs, .. } | ShapeContent::AutoShape { paragraphs, .. } => {
                paragraphs[0].runs.iter().map(|r| r.text.as_str()).collect()
            }
            other => panic!("expected a text shape, got {other:?}"),
        }
    }

    #[test]
    fn slidenum_field_prints_the_position_not_the_cache() {
        // The `<a:t>` inside a field is PowerPoint's cache of the last value it
        // printed; PowerPoint recomputes it. The probe deck states 777 on its
        // third slide and PowerPoint prints 3.
        let s = text_slide_with(
            r#"<a:r><a:rPr lang="en"/><a:t>p.</a:t></a:r>
               <a:fld id="{1}" type="slidenum"><a:rPr lang="en"/><a:t>777</a:t></a:fld>"#,
            3,
            1,
        );
        assert_eq!(only_text(&s), "p.3");
    }

    #[test]
    fn slidenum_field_counts_from_first_slide_num() {
        // firstSlideNum="5" on a six-slide deck prints 5..10 (COM probe).
        let s = text_slide_with(
            r#"<a:fld id="{1}" type="slidenum"><a:rPr lang="en"/><a:t>#</a:t></a:fld>"#,
            3,
            5,
        );
        assert_eq!(only_text(&s), "7");
    }

    #[test]
    fn slidenum_field_is_read_in_both_element_forms() {
        // A field with no cached text is self-closing and reaches quick-xml on
        // the EMPTY arm; a handler on the Start arm alone is silently inert.
        let s = text_slide_with(r#"<a:fld id="{1}" type="slidenum"/>"#, 12, 1);
        assert_eq!(only_text(&s), "12");
    }

    #[test]
    fn a_field_of_another_type_keeps_its_cached_text() {
        // Only `slidenum` is recomputed. A datetime field's cache is the last
        // string PowerPoint itself wrote, which is the best available answer.
        let s = text_slide_with(
            r#"<a:fld id="{1}" type="datetime1"><a:rPr lang="en"/><a:t>3/14/2026</a:t></a:fld>"#,
            2,
            1,
        );
        assert_eq!(only_text(&s), "3/14/2026");
    }

    #[test]
    fn test_detect_content_type() {
        assert_eq!(detect_content_type("image1.png"), Some("image/png".to_string()));
        assert_eq!(detect_content_type("photo.JPEG"), Some("image/jpeg".to_string()));
        assert_eq!(detect_content_type("logo.svg"), Some("image/svg+xml".to_string()));
        assert_eq!(detect_content_type("unknown.xyz"), None);
    }
}
