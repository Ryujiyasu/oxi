// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

use quick_xml::events::Event;
use quick_xml::reader::Reader;

use super::ParseError;
use super::theme::ThemeColors;
use crate::ir::{Alignment, ParagraphStyle, RunStyle, StyleDefinition, StyleSheet, TableStyle, TableConditionalFormat, CellBorders, BorderDef};

/// Resolve a theme font reference (e.g. "majorHAnsi", "minorEastAsia") to the actual font name.
/// OOXML theme values: majorHAnsi/minorHAnsi → Latin font, majorEastAsia/minorEastAsia → EA font.
fn resolve_theme_font(theme_val: &str, theme: &ThemeColors) -> Option<String> {
    resolve_theme_font_pub(theme_val, theme)
}

/// Public version of resolve_theme_font for use from ooxml.rs
pub fn resolve_theme_font_pub(theme_val: &str, theme: &ThemeColors) -> Option<String> {
    if theme_val.contains("EastAsia") {
        // "majorEastAsia" or "minorEastAsia" → use East Asian font
        // If theme has no EA font defined, default to MS Mincho (standard Japanese font)
        if theme_val.starts_with("major") {
            theme.major_font_ea.clone()
                .or_else(|| Some("MS Mincho".to_string()))
        } else {
            theme.minor_font_ea.clone()
                .or_else(|| Some("MS Mincho".to_string()))
        }
    } else if theme_val.contains("Bidi") && std::env::var("OXI_S987_DISABLE").is_err() {
        // S987 (Task educational__00235411): "majorBidi"/"minorBidi" resolve to
        // the theme's complex-script (Bidi) font — a non-empty <a:cs> else the
        // locale's Arab supplemental <a:font> — NOT the Latin major/minor. Word
        // renders educational__00235411's asciiTheme=majorBidi body in Times New
        // Roman (the Arab supplemental), not Cambria (Latin major), so its 12pt
        // double-spaced line pitch is 27.6pt not 28.14pt (+0.54pt/line → +1
        // page). Fall back to the Latin font when no Bidi font is derivable, so a
        // *Bidi token in a doc with no <a:cs>/Arab mapping stays byte-identical.
        // Opt-out OXI_S987_DISABLE (falls through to the Latin branch below).
        if theme_val.starts_with("major") {
            theme.major_font_bidi.clone().or_else(|| theme.major_font.clone())
        } else {
            theme.minor_font_bidi.clone().or_else(|| theme.minor_font.clone())
        }
    } else {
        // "majorHAnsi" / "minorHAnsi" → Latin font (and "*Bidi" when S987 is off)
        if theme_val.starts_with("major") {
            theme.major_font.clone()
        } else {
            theme.minor_font.clone()
        }
    }
}

pub fn parse_styles(xml: &str, theme: &ThemeColors) -> Result<StyleSheet, ParseError> {
    let mut reader = Reader::from_str(xml);
    let mut styles = StyleSheet::default();
    let mut in_doc_defaults = false;
    let mut in_rpr_default = false;
    let mut in_ppr_default = false;

    loop {
        match reader.read_event()? {
            Event::Start(e) => {
                let local = local_name(e.name().as_ref());
                match local.as_str() {
                    "docDefaults" => {
                        in_doc_defaults = true;
                    }
                    "rPrDefault" if in_doc_defaults => {
                        in_rpr_default = true;
                    }
                    "pPrDefault" if in_doc_defaults => {
                        in_ppr_default = true;
                    }
                    "rPr" if in_rpr_default => {
                        let run_style = parse_run_properties_block(&mut reader, theme)?;
                        styles.doc_default_run_style = Some(run_style);
                    }
                    "pPr" if in_ppr_default => {
                        let (para_style, align) = parse_para_properties_block_with_alignment(&mut reader)?;
                        styles.doc_default_para_style = Some(para_style);
                        styles.doc_default_alignment = align;
                    }
                    "style" => {
                        let mut style_id = None;
                        let mut style_type = None;
                        let mut is_default = false;
                        for attr in e.attributes().flatten() {
                            let key = local_name(attr.key.as_ref());
                            let val = String::from_utf8_lossy(&attr.value).to_string();
                            match key.as_str() {
                                "styleId" => style_id = Some(val),
                                "type" => style_type = Some(val),
                                "default" => is_default = val == "1",
                                _ => {}
                            }
                        }
                        if let (Some(id), Some(typ)) = (style_id, style_type) {
                            // Track default paragraph style ID
                            if typ == "paragraph" && is_default {
                                styles.default_paragraph_style_id = Some(id.clone());
                            }
                            // S871: track the default TABLE style too — a table
                            // with no w:tblStyle still inherits it (ECMA-376).
                            if typ == "table" && is_default {
                                styles.default_table_style_id = Some(id.clone());
                            }
                            if typ == "paragraph" || typ == "character" {
                                let (pstyle, based_on, align) = parse_style_definition(&mut reader, theme)?;
                                styles.styles.insert(
                                    id.clone(),
                                    StyleDefinition {
                                        style_id: id,
                                        based_on,
                                        paragraph: pstyle,
                                        alignment: align,
                                        resolved: false,
                                    },
                                );
                            } else if typ == "table" {
                                let (tbl_style, cond_fmts) = parse_table_style_definition(&mut reader)?;
                                if !cond_fmts.is_empty() {
                                    styles.table_conditional_formats.insert(id.clone(), cond_fmts);
                                }
                                styles.table_styles.insert(id, tbl_style);
                            }
                        }
                    }
                    _ => {}
                }
            }
            Event::End(e) => {
                let local = local_name(e.name().as_ref());
                match local.as_str() {
                    "docDefaults" => in_doc_defaults = false,
                    "rPrDefault" => in_rpr_default = false,
                    "pPrDefault" => in_ppr_default = false,
                    _ => {}
                }
            }
            Event::Eof => break,
            _ => {}
        }
    }

    // Resolve basedOn inheritance chains
    resolve_style_inheritance(&mut styles);

    // Resolve table style basedOn inheritance (e.g., Table Grid -> Normal Table)
    resolve_table_style_inheritance(&mut styles);

    Ok(styles)
}

/// Resolve basedOn inheritance: merge parent properties into child styles
fn resolve_style_inheritance(styles: &mut StyleSheet) {
    let ids: Vec<String> = styles.styles.keys().cloned().collect();
    for id in ids {
        resolve_single_style(styles, &id, 0);
    }
}

fn resolve_single_style(styles: &mut StyleSheet, id: &str, depth: u32) {
    // Prevent infinite recursion
    if depth > 20 {
        return;
    }

    // Check if already resolved or if style exists
    let based_on = {
        let Some(def) = styles.styles.get(id) else { return };
        if def.resolved {
            return;
        }
        def.based_on.clone()
    };

    // Resolve parent first
    if let Some(ref parent_id) = based_on {
        resolve_single_style(styles, parent_id, depth + 1);

        // Get parent's resolved properties
        let parent_para = styles.styles.get(parent_id).map(|d| d.paragraph.clone());
        // S539 (2026-06-11): jc lives on StyleDefinition.alignment (not
        // ParagraphStyle), so merge_para_style never propagated it down the
        // basedOn chain. 3a4f style "n" (no jc) basedOn "a"/Normal (jc=both)
        // resolved to None → paragraphs using "n" fell back to left and took
        // the non-justified break/render path, while Word justifies them.
        let parent_align = styles.styles.get(parent_id).and_then(|d| d.alignment);

        if let Some(parent) = parent_para {
            let def = styles.styles.get_mut(id).unwrap();
            // Merge: child overrides parent. Only fill in None/default fields from parent.
            merge_para_style(&mut def.paragraph, &parent);
            if def.alignment.is_none() {
                def.alignment = parent_align;
            }
        }
    }

    if let Some(def) = styles.styles.get_mut(id) {
        def.resolved = true;
    }
}

/// Merge parent paragraph style into child (child values take precedence)
fn merge_para_style(child: &mut ParagraphStyle, parent: &ParagraphStyle) {
    if child.heading_level.is_none() {
        child.heading_level = parent.heading_level;
    }
    if child.outline_level.is_none() {
        child.outline_level = parent.outline_level;
    }
    if child.line_spacing.is_none() {
        child.line_spacing = parent.line_spacing;
        child.line_spacing_rule = parent.line_spacing_rule.clone();
    }
    if child.space_before.is_none() {
        child.space_before = parent.space_before;
    }
    if child.space_after.is_none() {
        child.space_after = parent.space_after;
    }
    if child.before_lines.is_none() {
        child.before_lines = parent.before_lines;
    }
    if child.after_lines.is_none() {
        child.after_lines = parent.after_lines;
    }
    // S800 (2026-07-12, opt-out OXI_S800_DISABLE): a style that DECLARES its own
    // numPr does NOT inherit ANCESTOR left/firstLine indents — Word inserts the
    // numbering level's ind at the numPr-declaring layer, overriding ind
    // inherited from ABOVE it. ukframework FWKBullet/FWKBodyText: numPr in the
    // style, ind left=720 inherited from the ListParagraph ancestor, yet Word
    // renders at the LEVEL ind (1436/792 — PDF text x = margin+71.8/39.6). A
    // style that declares BOTH numPr and its OWN ind keeps that ind (nyserda
    // ListBullet2 ind left=1152 beats level 5490 = the S771 case, unchanged —
    // the level ind is weaker than the declaring layer's own props but stronger
    // than the ancestors'). With the inherit cut, the resolved style's
    // indent_left is None and the existing S771 gate in ooxml.rs
    // (style.indent_left.is_none()) applies the level ind. numId="0"
    // (numbering REMOVAL) keeps normal inheritance — no level ind will apply.
    let s800_numpr_layer = child.num_id.as_deref().map_or(false, |n| !n.is_empty() && n != "0")
        && std::env::var("OXI_S800_DISABLE").is_err();
    if child.indent_left.is_none() && !s800_numpr_layer {
        child.indent_left = parent.indent_left;
    }
    if child.indent_right.is_none() {
        child.indent_right = parent.indent_right;
    }
    if child.indent_first_line.is_none() && !s800_numpr_layer {
        child.indent_first_line = parent.indent_first_line;
    }
    if child.num_id.is_none() {
        child.num_id = parent.num_id.clone();
        if child.num_id.is_some() {
            child.num_ilvl = parent.num_ilvl;
        }
    }
    // Inherit bool fields: true from parent if child doesn't set
    // S955: three-state keepNext/keepLines — a child style's explicit
    // `w:val="0"` must survive a basedOn parent's ON (legal__0010437a:
    // MiscellaneousBody keepNext=0 basedOn MiscellaneousHeading keepNext=ON;
    // the monotone merge turned every body paragraph into a keepNext chain →
    // repeated chain pushes → footer-only phantom pages). The
    // has_explicit_widow_control pattern.
    // Probe-proven Word-correct (kn0_probe: Word honors a DIRECT val=0 even
    // against a style's explicit val=1). Shipped together with S956, which
    // removed the one blocker: mysignaiguide's PASS rode the WRONG keepNext,
    // compensating an empty-paragraph mark-font error. Opt-out
    // OXI_S955_DISABLE.
    if std::env::var("OXI_S955_DISABLE").is_err() {
        if !child.has_explicit_keep_next {
            if parent.has_explicit_keep_next {
                child.keep_next = parent.keep_next;
                child.has_explicit_keep_next = true;
            } else if parent.keep_next {
                child.keep_next = true;
            }
        }
        if !child.has_explicit_keep_lines {
            if parent.has_explicit_keep_lines {
                child.keep_lines = parent.keep_lines;
                child.has_explicit_keep_lines = true;
            } else if parent.keep_lines {
                child.keep_lines = true;
            }
        }
    } else {
        if !child.keep_next && parent.keep_next {
            child.keep_next = true;
        }
        if !child.keep_lines && parent.keep_lines {
            child.keep_lines = true;
        }
    }
    if !child.has_explicit_widow_control && parent.has_explicit_widow_control {
        child.widow_control = parent.widow_control;
        child.has_explicit_widow_control = true;
    }
    // S985: CT_OnOff three-state (mirrors keepNext/keepLines/widowControl above).
    // A child style that EXPLICITLY set pageBreakBefore (incl. val="0") keeps its
    // own value; only a child that did NOT set it inherits — the parent's value +
    // explicitness if the parent set it, else the legacy monotonic OR. This stops
    // a derived style's explicit `val="0"` from being overwritten by a basedOn
    // parent's true (legal__001410a8 yHeading2/nHeading2).
    if std::env::var("OXI_S985_DISABLE").is_err() {
        if !child.has_explicit_page_break_before {
            if parent.has_explicit_page_break_before {
                child.page_break_before = parent.page_break_before;
                child.has_explicit_page_break_before = true;
            } else if parent.page_break_before {
                child.page_break_before = true;
            }
        }
    } else if !child.page_break_before && parent.page_break_before {
        child.page_break_before = true;
    }
    if !child.contextual_spacing && parent.contextual_spacing {
        child.contextual_spacing = true;
    }
    // S675: before/afterAutospacing is NOT inherited — Word applies it only when
    // set directly on the paragraph (COM-confirmed harassbosi/b837 Web style = 0).
    if !child.bidi && parent.bidi {
        child.bidi = true;
    }
    // Inherit snap_to_grid (parent false overrides child default true)
    if !parent.snap_to_grid {
        child.snap_to_grid = false;
    }
    if child.shading.is_none() {
        child.shading = parent.shading.clone();
    }
    if child.borders.is_none() {
        child.borders = parent.borders.clone();
    }
    if std::env::var("OXI_S977_DISABLE").is_err() {
        // S977: accumulate through the basedOn chain instead of replacing.
        child.tab_stops = merge_tab_stops(&child.tab_stops, &parent.tab_stops);
    } else if child.tab_stops.is_empty() {
        child.tab_stops = parent.tab_stops.clone();
    }
    if child.default_run_style.is_none() {
        child.default_run_style = parent.default_run_style.clone();
    } else if let Some(ref parent_rs) = parent.default_run_style {
        // Merge run style: child overrides parent
        let child_rs = child.default_run_style.as_mut().unwrap();
        merge_run_style(child_rs, parent_rs);
    }
}

/// Merge parent run style into child (child values take precedence)
pub(crate) fn merge_run_style(child: &mut RunStyle, parent: &RunStyle) {
    if child.font_family.is_none() {
        child.font_family = parent.font_family.clone();
    }
    if child.font_family_east_asia.is_none() {
        child.font_family_east_asia = parent.font_family_east_asia.clone();
    }
    // §4.6.3: explicit eastAsia attribute is inherited (sticky once set anywhere
    // up the chain). Theme fallback never sets it.
    if !child.has_explicit_east_asia && parent.has_explicit_east_asia {
        child.has_explicit_east_asia = true;
    }
    if child.font_size.is_none() {
        child.font_size = parent.font_size;
    }
    if child.color.is_none() {
        child.color = parent.color.clone();
    }
    if child.highlight.is_none() {
        child.highlight = parent.highlight.clone();
    }
    if child.character_spacing.is_none() {
        child.character_spacing = parent.character_spacing;
    }
    if child.shading.is_none() {
        child.shading = parent.shading.clone();
    }
    if child.font_size_cs.is_none() {
        child.font_size_cs = parent.font_size_cs;
    }
    if child.kern.is_none() {
        child.kern = parent.kern;
    }
    if child.text_scale.is_none() {
        child.text_scale = parent.text_scale;
    }
    // Boolean properties: inherit from parent if child doesn't explicitly set them.
    // S976: bold/italic are three-state — a run that carries `<w:b w:val="0"/>`
    // has explicitly turned bold OFF and must not pick the parent's ON back up
    // (Word then measures it with REGULAR glyph widths; the monotonic rule
    // measured a non-bold run of a bold heading style ~10% too wide).
    if std::env::var("OXI_S976_DISABLE").is_err() {
        if !child.has_explicit_bold {
            if parent.has_explicit_bold {
                child.bold = parent.bold;
                child.has_explicit_bold = true;
            } else if parent.bold {
                child.bold = true;
            }
        }
        if !child.has_explicit_italic {
            if parent.has_explicit_italic {
                child.italic = parent.italic;
                child.has_explicit_italic = true;
            } else if parent.italic {
                child.italic = true;
            }
        }
    } else {
        if !child.bold && parent.bold {
            child.bold = true;
        }
        if !child.italic && parent.italic {
            child.italic = true;
        }
    }
    if !child.underline && parent.underline {
        child.underline = true;
    }
    // S988A: strikethrough is three-state (mirror bold/italic above) — an
    // explicit child `<w:strike w:val="false"/>` beats the parent's ON, while a
    // non-explicit child still inherits via the same monotonic OR as before.
    // Opt-out OXI_S988A_DISABLE keeps the legacy monotonic OR. NOTE: caps/
    // dstrike/smallCaps deliberately keep their pre-S988A no-inheritance in the
    // style chain — the report's verified A/B removed val="false" NODES (= the
    // producer fix below), it did NOT test adding caps INHERITANCE (which would
    // newly fire on ~129 golden docs with basedOn+plain-caps). The target
    // (CSILevel, no basedOn) needs only the producer's val="false" honoring;
    // caps-through-basedOn is a separate, unverified correctness change.
    if std::env::var("OXI_S988A_DISABLE").is_err() {
        if !child.has_explicit_strikethrough {
            if parent.has_explicit_strikethrough {
                child.strikethrough = parent.strikethrough;
                child.has_explicit_strikethrough = true;
            } else if parent.strikethrough {
                child.strikethrough = true;
            }
        }
    } else if !child.strikethrough && parent.strikethrough {
        child.strikethrough = true;
    }
    // Round 29: vertical_align inheritance (footnote reference style "aa"
    // sets vertAlign=superscript via the character style chain)
    if child.vertical_align.is_none() && parent.vertical_align.is_some() {
        child.vertical_align = parent.vertical_align;
    }
}

/// Parse a run properties block (<w:rPr>...</w:rPr>) and return RunStyle
fn parse_run_properties_block(reader: &mut Reader<&[u8]>, theme: &ThemeColors) -> Result<RunStyle, ParseError> {
    let mut rs = RunStyle::default();
    let mut depth = 1;

    loop {
        match reader.read_event()? {
            Event::Start(e) => {
                let local = local_name(e.name().as_ref());
                if local == "rFonts" {
                    for attr in e.attributes().flatten() {
                        let key = local_name(attr.key.as_ref());
                        if key == "ascii" || key == "hAnsi" {
                            rs.font_family =
                                Some(String::from_utf8_lossy(&attr.value).to_string());
                        } else if key == "asciiTheme" || key == "hAnsiTheme" {
                            if rs.font_family.is_none() {
                                let val = String::from_utf8_lossy(&attr.value);
                                let font = resolve_theme_font(&val, theme);
                                if let Some(f) = font {
                                    rs.font_family = Some(f);
                                }
                            }
                        } else if key == "eastAsia" {
                            rs.font_family_east_asia =
                                Some(String::from_utf8_lossy(&attr.value).to_string());
                            rs.has_explicit_east_asia = true;
                        } else if key == "eastAsiaTheme" {
                            if rs.font_family_east_asia.is_none() {
                                let val = String::from_utf8_lossy(&attr.value);
                                let font = if val.starts_with("major") {
                                    theme.major_font_ea.clone().or_else(|| theme.major_font.clone())
                                } else {
                                    theme.minor_font_ea.clone().or_else(|| theme.minor_font.clone())
                                };
                                if let Some(f) = font {
                                    rs.font_family_east_asia = Some(f);
                                }
                            }
                        }
                    }
                }
                depth += 1;
            }
            Event::Empty(e) => {
                apply_run_property_empty(&e, &mut rs, theme);
            }
            Event::End(_) => {
                depth -= 1;
                if depth == 0 {
                    break;
                }
            }
            Event::Eof => break,
            _ => {}
        }
    }

    Ok(rs)
}

/// Parse pPrDefault's pPr block, extracting both ParagraphStyle and alignment (jc).
/// Regular parse_para_properties_block doesn't capture jc since ParagraphStyle has no alignment field.
fn parse_para_properties_block_with_alignment(reader: &mut Reader<&[u8]>) -> Result<(ParagraphStyle, Option<Alignment>), ParseError> {
    let mut style = ParagraphStyle::default();
    let mut alignment: Option<Alignment> = None;
    let mut depth = 1;

    loop {
        match reader.read_event()? {
            Event::Start(e) => {
                let local = local_name(e.name().as_ref());
                if local == "pBdr" && depth == 1 {
                    style.borders = Some(super::ooxml::parse_paragraph_borders(reader)?);
                } else {
                    depth += 1;
                }
            }
            Event::Empty(e) => {
                let local = local_name(e.name().as_ref());
                if local == "jc" {
                    for attr in e.attributes().flatten() {
                        if local_name(attr.key.as_ref()) == "val" {
                            let val = String::from_utf8_lossy(&attr.value);
                            alignment = Some(match val.as_ref() {
                                "left" | "start" => Alignment::Left,
                                "center" => Alignment::Center,
                                "right" | "end" => Alignment::Right,
                                "both" => Alignment::Justify,
                                "distribute" => Alignment::Distribute,
                                _ => Alignment::Left,
                            });
                        }
                    }
                } else {
                    apply_para_property_empty(&e, &mut style);
                }
            }
            Event::End(_) => {
                depth -= 1;
                if depth == 0 { break; }
            }
            Event::Eof => break,
            _ => {}
        }
    }

    Ok((style, alignment))
}

/// Parse a paragraph properties block (<w:pPr>...</w:pPr>) and return ParagraphStyle
#[allow(dead_code)]
fn parse_para_properties_block(reader: &mut Reader<&[u8]>) -> Result<ParagraphStyle, ParseError> {
    let mut style = ParagraphStyle::default();
    let mut depth = 1;

    loop {
        match reader.read_event()? {
            Event::Start(e) => {
                let local = local_name(e.name().as_ref());
                if local == "pBdr" && depth == 1 {
                    style.borders = Some(super::ooxml::parse_paragraph_borders(reader)?);
                } else {
                    depth += 1;
                }
            }
            Event::Empty(e) => {
                apply_para_property_empty(&e, &mut style);
            }
            Event::End(_) => {
                depth -= 1;
                if depth == 0 {
                    break;
                }
            }
            Event::Eof => break,
            _ => {}
        }
    }

    Ok(style)
}

/// S977: merge a child's tab stops into an inherited list. ECMA-376
/// §17.3.1.38: custom tab stops ACCUMULATE through the style hierarchy (they do
/// not replace), and `<w:tab w:val="clear" w:pos="N"/>` removes the inherited
/// stop at N. The child's entry wins at an equal position. Clear directives
/// never survive, so layout only ever sees real stops.
pub(crate) fn merge_tab_stops(child: &[crate::ir::TabStop], parent: &[crate::ir::TabStop]) -> Vec<crate::ir::TabStop> {
    let same = |a: f32, b: f32| (a - b).abs() < 0.05;
    let mut out: Vec<crate::ir::TabStop> = parent.iter().filter(|t| !t.clear).cloned().collect();
    for c in child {
        out.retain(|t| !same(t.position, c.position));
        if !c.clear {
            out.push(c.clone());
        }
    }
    out.sort_by(|a, b| {
        a.position
            .partial_cmp(&b.position)
            .unwrap_or(std::cmp::Ordering::Equal)
    });
    out
}

/// S976: read a CT_OnOff element's value. The element being present IS the
/// setting; `w:val` only says which way. Absent val means ON.
pub(crate) fn ct_on_off(e: &quick_xml::events::BytesStart) -> bool {
    for attr in e.attributes().flatten() {
        if local_name(attr.key.as_ref()) == "val" {
            let val = String::from_utf8_lossy(&attr.value);
            return val.as_ref() != "0" && val.as_ref() != "false" && val.as_ref() != "off";
        }
    }
    true
}

/// Apply a single empty run property element to RunStyle
fn apply_run_property_empty(e: &quick_xml::events::BytesStart, rs: &mut RunStyle, theme: &ThemeColors) {
    let local = local_name(e.name().as_ref());
    match local.as_str() {
        "b" => {
            rs.bold = ct_on_off(e);
            rs.has_explicit_bold = true;
        }
        "i" => {
            rs.italic = ct_on_off(e);
            rs.has_explicit_italic = true;
        }
        "u" => {
            rs.underline = true;
            for attr in e.attributes().flatten() {
                if local_name(attr.key.as_ref()) == "val" {
                    let val = String::from_utf8_lossy(&attr.value).to_string();
                    if val != "none" {
                        rs.underline_style = Some(val);
                    } else {
                        rs.underline = false;
                    }
                }
            }
        }
        "strike" => {
            // S988A: honor `w:val="false"` (CT_OnOff three-state, mirrors S976
            // bold/italic). Opt-out OXI_S988A_DISABLE keeps the presence-only ON.
            if std::env::var("OXI_S988A_DISABLE").is_err() {
                rs.strikethrough = ct_on_off(e);
                rs.has_explicit_strikethrough = true;
            } else {
                rs.strikethrough = true;
            }
        }
        "dstrike" => {
            if std::env::var("OXI_S988A_DISABLE").is_err() {
                let v = ct_on_off(e);
                rs.strikethrough = v;
                rs.double_strikethrough = v;
                rs.has_explicit_strikethrough = true;
                rs.has_explicit_double_strikethrough = true;
            } else {
                rs.strikethrough = true;
                rs.double_strikethrough = true;
            }
        }
        "sz" => {
            for attr in e.attributes().flatten() {
                if local_name(attr.key.as_ref()) == "val" {
                    let val = String::from_utf8_lossy(&attr.value);
                    rs.font_size = val.parse::<f32>().ok().map(|v| v / 2.0);
                }
            }
        }
        "color" => {
            let mut color_val = None;
            let mut theme_color = None;
            let mut theme_tint = None;
            let mut theme_shade = None;
            for attr in e.attributes().flatten() {
                let key = local_name(attr.key.as_ref());
                let val = String::from_utf8_lossy(&attr.value).to_string();
                match key.as_str() {
                    "val" => color_val = Some(val),
                    "themeColor" => theme_color = Some(val),
                    "themeTint" => theme_tint = Some(val),
                    "themeShade" => theme_shade = Some(val),
                    _ => {}
                }
            }
            if let Some(ref tc) = theme_color {
                if let Some(resolved) = theme.resolve(tc) {
                    let mut hex = resolved.clone();
                    if let Some(ref tint) = theme_tint {
                        if let Ok(t) = u8::from_str_radix(tint, 16) {
                            hex = ThemeColors::apply_theme_tint(&hex, t);
                        }
                    }
                    if let Some(ref shade) = theme_shade {
                        if let Ok(s) = u8::from_str_radix(shade, 16) {
                            let shade_val = -(1.0 - s as f64 / 255.0);
                            hex = ThemeColors::apply_tint_shade(&hex, shade_val);
                        }
                    }
                    rs.color = Some(hex);
                } else if let Some(ref cv) = color_val {
                    if cv != "auto" { rs.color = Some(cv.clone()); }
                }
            } else if let Some(ref cv) = color_val {
                if cv != "auto" { rs.color = Some(cv.clone()); }
            }
        }
        // TODO: rFonts theme attributes (asciiTheme, hAnsiTheme, eastAsiaTheme) are
        // resolved via ThemeColors in parse_run_properties_block, but this function
        // receives ThemeColors so they are handled here. However, the empty-element
        // rFonts in parse_style_definition's rPr branch does NOT have ThemeColors
        // access — see the TODO there for that limitation.
        "rFonts" => {
            for attr in e.attributes().flatten() {
                let key = local_name(attr.key.as_ref());
                if key == "ascii" || key == "hAnsi" {
                    rs.font_family =
                        Some(String::from_utf8_lossy(&attr.value).to_string());
                } else if key == "asciiTheme" || key == "hAnsiTheme" {
                    if rs.font_family.is_none() {
                        let val = String::from_utf8_lossy(&attr.value);
                        let font = if val.starts_with("major") {
                            theme.major_font.clone()
                        } else {
                            theme.minor_font.clone()
                        };
                        if let Some(f) = font {
                            rs.font_family = Some(f);
                        }
                    }
                } else if key == "eastAsia" {
                    rs.font_family_east_asia =
                        Some(String::from_utf8_lossy(&attr.value).to_string());
                    rs.has_explicit_east_asia = true;
                } else if key == "eastAsiaTheme" {
                    if rs.font_family_east_asia.is_none() {
                        let val = String::from_utf8_lossy(&attr.value);
                        let font = if val.starts_with("major") {
                            theme.major_font_ea.clone().or_else(|| theme.major_font.clone())
                        } else {
                            theme.minor_font_ea.clone().or_else(|| theme.minor_font.clone())
                        };
                        if let Some(f) = font {
                            rs.font_family_east_asia = Some(f);
                        }
                    }
                }
            }
        }
        "lang" => {
            // S763c: capture the East-Asian language (w:lang w:eastAsia) — drives
            // Word's ambiguous curly-quote font choice (CJK lang → eastAsia font,
            // Latin lang → Latin font).
            // S956: also capture the LATIN language (w:lang w:val). A CJK value
            // there (the unusual `w:val="ja"`) makes Word resolve even Latin
            // text — and the ¶ mark — through the East Asian font chain.
            for attr in e.attributes().flatten() {
                match local_name(attr.key.as_ref()).as_str() {
                    "eastAsia" => {
                        rs.east_asia_lang = Some(String::from_utf8_lossy(&attr.value).to_string());
                    }
                    "val" => {
                        rs.latin_lang = Some(String::from_utf8_lossy(&attr.value).to_string());
                    }
                    _ => {}
                }
            }
        }
        "spacing" => {
            for attr in e.attributes().flatten() {
                if local_name(attr.key.as_ref()) == "val" {
                    let val = String::from_utf8_lossy(&attr.value);
                    rs.character_spacing = val.parse::<f32>().ok().map(|v| v / 20.0);
                }
            }
        }
        "smallCaps" => {
            if std::env::var("OXI_S988A_DISABLE").is_err() {
                rs.small_caps = ct_on_off(e);
                rs.has_explicit_small_caps = true;
            } else {
                rs.small_caps = true;
            }
        }
        "caps" => {
            if std::env::var("OXI_S988A_DISABLE").is_err() {
                rs.all_caps = ct_on_off(e);
                rs.has_explicit_all_caps = true;
            } else {
                rs.all_caps = true;
            }
        }
        "shd" => {
            for attr in e.attributes().flatten() {
                let key = local_name(attr.key.as_ref());
                if key == "fill" {
                    let val = String::from_utf8_lossy(&attr.value).to_string();
                    if val != "auto" {
                        rs.shading = Some(val);
                    }
                }
            }
        }
        "rtl" => rs.rtl = true,
        "vanish" => rs.vanish = true,
        // w:webHidden is Web-Layout-only; print/PDF renders it (ToC page numbers).
        "webHidden" => {}
        "outline" => rs.outline = true,
        "shadow" => rs.shadow = true,
        "emboss" => rs.emboss = true,
        "imprint" => rs.imprint = true,
        "bCs" => rs.bold_cs = true,
        "iCs" => rs.italic_cs = true,
        "szCs" => {
            for attr in e.attributes().flatten() {
                if local_name(attr.key.as_ref()) == "val" {
                    let val = String::from_utf8_lossy(&attr.value);
                    rs.font_size_cs = val.parse::<f32>().ok().map(|v| v / 2.0);
                }
            }
        }
        "kern" => {
            for attr in e.attributes().flatten() {
                if local_name(attr.key.as_ref()) == "val" {
                    let val = String::from_utf8_lossy(&attr.value);
                    rs.kern = val.parse::<f32>().ok().map(|v| v / 2.0);
                }
            }
        }
        "fitText" => {
            for attr in e.attributes().flatten() {
                let ln = local_name(attr.key.as_ref());
                if ln == "val" {
                    let val = String::from_utf8_lossy(&attr.value);
                    rs.fit_text = val.parse::<f32>().ok().map(|v| v / 20.0);
                } else if ln == "id" {
                    let val = String::from_utf8_lossy(&attr.value);
                    rs.fit_text_id = val.parse::<i64>().ok();
                }
            }
        }
        "w" => {
            for attr in e.attributes().flatten() {
                if local_name(attr.key.as_ref()) == "val" {
                    let val = String::from_utf8_lossy(&attr.value);
                    rs.text_scale = val.parse::<f32>().ok();
                }
            }
        }
        "eastAsianLayout" => {
            for attr in e.attributes().flatten() {
                let key = local_name(attr.key.as_ref());
                match key.as_str() {
                    "combine" => {
                        let val = String::from_utf8_lossy(&attr.value);
                        rs.combine = val.as_ref() != "0" && val.as_ref() != "false";
                    }
                    "vert" => {
                        let val = String::from_utf8_lossy(&attr.value);
                        rs.vert_in_horz = val.as_ref() != "0" && val.as_ref() != "false";
                    }
                    _ => {}
                }
            }
        }
        // Round 29 (2026-04-08): vertAlign for character styles. The
        // "footnote reference" character style (id "aa" in many docs) sets
        // <w:vertAlign w:val="superscript"/> — without this parser branch,
        // the rStyle inheritance never picks it up and footnote markers
        // render at body baseline/size instead of small superscript.
        "vertAlign" => {
            for attr in e.attributes().flatten() {
                if local_name(attr.key.as_ref()) == "val" {
                    let val = String::from_utf8_lossy(&attr.value);
                    rs.vertical_align = match val.as_ref() {
                        "superscript" => Some(crate::ir::VerticalAlign::Superscript),
                        "subscript" => Some(crate::ir::VerticalAlign::Subscript),
                        _ => None,
                    };
                }
            }
        }
        _ => {}
    }
}

/// Apply a single empty paragraph property element to ParagraphStyle
fn apply_para_property_empty(e: &quick_xml::events::BytesStart, style: &mut ParagraphStyle) {
    let local = local_name(e.name().as_ref());
    match local.as_str() {
        "spacing" => {
            let mut line_val: Option<f32> = None;
            let mut line_rule: Option<String> = None;
            for attr in e.attributes().flatten() {
                let key = local_name(attr.key.as_ref());
                let val = String::from_utf8_lossy(&attr.value);
                match key.as_str() {
                    "before" => {
                        style.space_before = val.parse::<f32>().ok().map(|v| v / 20.0);
                    }
                    "after" => {
                        style.space_after = val.parse::<f32>().ok().map(|v| v / 20.0);
                    }
                    "beforeLines" => {
                        style.before_lines = val.parse::<f32>().ok();
                    }
                    "afterLines" => {
                        style.after_lines = val.parse::<f32>().ok();
                    }
                    // NOTE: before/afterAutospacing is INTENTIONALLY NOT parsed at the
                    // style level — Word applies HTML-paragraph autospacing only when the
                    // attribute is set DIRECTLY on the paragraph's pPr, not when inherited
                    // from a paragraph style (S675, COM-confirmed: harassbosi/b837 64 Web
                    // paras render with 0 extra; a direct-pPr para renders 13.75).
                    // S864: retain style-level HTML autospacing for the table-cell
                    // parser. Body paragraphs still intentionally ignore it (S675).
                    "beforeAutospacing" if crate::layout::s864_part("F") => {
                        style.before_autospacing = val == "1" || val == "true" || val == "on";
                    }
                    "afterAutospacing" if crate::layout::s864_part("F") => {
                        style.after_autospacing = val == "1" || val == "true" || val == "on";
                    }
                    "line" => {
                        line_val = val.parse::<f32>().ok();
                    }
                    "lineRule" => {
                        line_rule = Some(val.to_string());
                    }
                    _ => {}
                }
            }
            if let Some(lv) = line_val {
                match line_rule.as_deref() {
                    Some("exact") => {
                        style.line_spacing = Some(lv / 20.0);
                        style.line_spacing_rule = Some("exact".to_string());
                    }
                    Some("atLeast") => {
                        style.line_spacing = Some(lv / 20.0);
                        style.line_spacing_rule = Some("atLeast".to_string());
                    }
                    _ if lv < 0.0 => {
                        // R7.72: negative w:line with lineRule="auto" → exact mode |val|/20
                        // COM-confirmed 2026-05-15 on d4d126.
                        style.line_spacing = Some(lv.abs() / 20.0);
                        style.line_spacing_rule = Some("exact".to_string());
                    }
                    _ => {
                        style.line_spacing = Some(lv / 240.0);
                    }
                }
            }
        }
        "ind" => {
            for attr in e.attributes().flatten() {
                let key = local_name(attr.key.as_ref());
                let val = String::from_utf8_lossy(&attr.value);
                match key.as_str() {
                    "left" => {
                        style.indent_left =
                            val.parse::<f32>().ok().map(|v| v / 20.0);
                    }
                    "right" => {
                        style.indent_right =
                            val.parse::<f32>().ok().map(|v| v / 20.0);
                    }
                    "firstLine" => {
                        style.indent_first_line =
                            val.parse::<f32>().ok().map(|v| v / 20.0);
                    }
                    "hanging" => {
                        // Hanging indent: negative first-line indent
                        style.indent_first_line =
                            val.parse::<f32>().ok().map(|v| -(v / 20.0));
                    }
                    _ => {}
                }
            }
        }
        "shd" => {
            for attr in e.attributes().flatten() {
                if local_name(attr.key.as_ref()) == "fill" {
                    let val = String::from_utf8_lossy(&attr.value).to_string();
                    if val != "auto" {
                        style.shading = Some(val);
                    }
                }
            }
        }
        "snapToGrid" => {
            // Presence alone means true; explicit val="0"/"false" disables
            let mut enabled = true;
            for attr in e.attributes().flatten() {
                if local_name(attr.key.as_ref()) == "val" {
                    let val = String::from_utf8_lossy(&attr.value);
                    enabled = val.as_ref() != "0" && val.as_ref() != "false";
                }
            }
            style.snap_to_grid = enabled;
        }
        "autoSpaceDE" => {
            // Session 85 fix: parse autoSpaceDE in STYLE definitions (was missing).
            // tokumei_08_01 series uses style "ac" with <w:autoSpaceDE w:val="0"/>;
            // without this parse, the style stored auto_space_de=true (default) and
            // paragraph-level pStyle inheritance silently used true. Result: 4 ASCII/CJK
            // boundary chars × 2.5pt = 10pt extra per line → 36-char paragraphs wrap to
            // 35+1 in Oxi vs 1-line in Word (a1d6 p4 direction-flip cause).
            let mut enabled = true;
            for attr in e.attributes().flatten() {
                if local_name(attr.key.as_ref()) == "val" {
                    let val = String::from_utf8_lossy(&attr.value);
                    enabled = val.as_ref() != "0" && val.as_ref() != "false";
                }
            }
            style.auto_space_de = enabled;
        }
        "autoSpaceDN" => {
            let mut enabled = true;
            for attr in e.attributes().flatten() {
                if local_name(attr.key.as_ref()) == "val" {
                    let val = String::from_utf8_lossy(&attr.value);
                    enabled = val.as_ref() != "0" && val.as_ref() != "false";
                }
            }
            style.auto_space_dn = enabled;
        }
        "wordWrap" => {
            let mut enabled = true;
            for attr in e.attributes().flatten() {
                if local_name(attr.key.as_ref()) == "val" {
                    let val = String::from_utf8_lossy(&attr.value);
                    enabled = val.as_ref() != "0" && val.as_ref() != "false";
                }
            }
            style.word_wrap = enabled;
        }
        "pageBreakBefore" => {
            // CT_OnOff: w:val="0"/"false"/"off" disables (the Heading styles in
            // the AI-guideline corpus carry <w:pageBreakBefore w:val="0"/> — Word
            // honours val=0 and does NOT break; the old presence-only parse forced
            // a break before every styled heading → 3-page docs blew up to 9-17
            // pages). See S597.
            let mut enabled = true;
            for attr in e.attributes().flatten() {
                if local_name(attr.key.as_ref()) == "val" {
                    let val = String::from_utf8_lossy(&attr.value);
                    enabled = val.as_ref() != "0" && val.as_ref() != "false" && val.as_ref() != "off";
                }
            }
            style.page_break_before = enabled;
            // S985 (Task legal__001410a8, 2026-07-22): record explicitness so a
            // derived style's `<w:pageBreakBefore w:val="0"/>` beats a basedOn
            // parent's ON (the CT_OnOff three-state — S633 keepNext / S976 bold /
            // S606b snapToGrid class; S884 fixed the DIRECT pPr path, this fixes
            // the STYLE chain). Without it merge_paragraph_style's monotonic OR
            // restores the parent's true (legal__001410a8 yHeading2/nHeading2 →
            // spurious page breaks before Forms / Details of applicant → +2 pages).
            // Opt-out OXI_S985_DISABLE (gates producer + merge → legacy monotonic OR).
            if std::env::var("OXI_S985_DISABLE").is_err() {
                style.has_explicit_page_break_before = true;
            }
        }
        "contextualSpacing" => {
            let mut enabled = true;
            for attr in e.attributes().flatten() {
                if local_name(attr.key.as_ref()) == "val" {
                    let val = String::from_utf8_lossy(&attr.value);
                    enabled = val.as_ref() != "0" && val.as_ref() != "false";
                }
            }
            style.contextual_spacing = enabled;
        }
        "keepNext" => {
            // CT_OnOff: val="0"/"false"/"off" disables (S633).
            let mut enabled = true;
            for attr in e.attributes().flatten() {
                if local_name(attr.key.as_ref()) == "val" {
                    let val = String::from_utf8_lossy(&attr.value);
                    enabled = val.as_ref() != "0" && val.as_ref() != "false" && val.as_ref() != "off";
                }
            }
            style.keep_next = enabled;
            style.has_explicit_keep_next = true;
        }
        "keepLines" => {
            let mut enabled = true;
            for attr in e.attributes().flatten() {
                if local_name(attr.key.as_ref()) == "val" {
                    let val = String::from_utf8_lossy(&attr.value);
                    enabled = val.as_ref() != "0" && val.as_ref() != "false" && val.as_ref() != "off";
                }
            }
            style.keep_lines = enabled;
            style.has_explicit_keep_lines = true;
        }
        "widowControl" => {
            let mut enabled = true;
            for attr in e.attributes().flatten() {
                if local_name(attr.key.as_ref()) == "val" {
                    let val = String::from_utf8_lossy(&attr.value);
                    enabled = val.as_ref() != "0" && val.as_ref() != "false";
                }
            }
            style.widow_control = enabled;
            style.has_explicit_widow_control = true;
        }
        "bidi" => {
            let mut enabled = true;
            for attr in e.attributes().flatten() {
                if local_name(attr.key.as_ref()) == "val" {
                    let val = String::from_utf8_lossy(&attr.value);
                    enabled = val.as_ref() != "0" && val.as_ref() != "false";
                }
            }
            style.bidi = enabled;
        }
        "textAlignment" => {
            for attr in e.attributes().flatten() {
                if local_name(attr.key.as_ref()) == "val" {
                    let val = String::from_utf8_lossy(&attr.value).to_string();
                    style.text_alignment = Some(val);
                }
            }
        }
        _ => {}
    }
}

/// Parse a style definition block, returning (ParagraphStyle, Option<basedOn_id>)
fn parse_style_definition(
    reader: &mut Reader<&[u8]>,
    theme: &ThemeColors,
) -> Result<(ParagraphStyle, Option<String>, Option<Alignment>), ParseError> {
    let mut style = ParagraphStyle::default();
    let mut run_style = RunStyle::default();
    let mut has_run_style = false;
    let mut based_on: Option<String> = None;
    let mut alignment: Option<Alignment> = None;
    let mut depth = 0;
    let mut in_rpr = false;
    let mut in_num_pr = false;

    loop {
        match reader.read_event()? {
            Event::Start(e) => {
                let local = local_name(e.name().as_ref());
                match local.as_str() {
                    "pBdr" if !in_rpr && !in_num_pr => {
                        style.borders = Some(super::ooxml::parse_paragraph_borders(reader)?);
                        continue;
                    }
                    "numPr" if !in_rpr => {
                        in_num_pr = true;
                        depth += 1;
                    }
                    "tabs" if !in_rpr && !in_num_pr => {
                        if let Ok(stops) = super::ooxml::parse_tab_stops(reader) {
                            style.tab_stops = stops;
                        }
                        // parse_tab_stops consumes up to </tabs>, don't increment depth
                        continue;
                    }
                    "rPr" if depth == 0 => {
                        in_rpr = true;
                        depth += 1;
                    }
                    "rFonts" if in_rpr => {
                        for attr in e.attributes().flatten() {
                            let key = local_name(attr.key.as_ref());
                            if key == "ascii" || key == "hAnsi" {
                                run_style.font_family =
                                    Some(String::from_utf8_lossy(&attr.value).to_string());
                                has_run_style = true;
                            } else if key == "asciiTheme" || key == "hAnsiTheme" {
                                if run_style.font_family.is_none() {
                                    let val = String::from_utf8_lossy(&attr.value);
                                    let font = if val.starts_with("major") {
                                        theme.major_font.clone()
                                    } else {
                                        theme.minor_font.clone()
                                    };
                                    if let Some(f) = font {
                                        run_style.font_family = Some(f);
                                        has_run_style = true;
                                    }
                                }
                            } else if key == "eastAsia" {
                                run_style.font_family_east_asia =
                                    Some(String::from_utf8_lossy(&attr.value).to_string());
                                run_style.has_explicit_east_asia = true;
                                has_run_style = true;
                            } else if key == "eastAsiaTheme" {
                                if run_style.font_family_east_asia.is_none() {
                                    let val = String::from_utf8_lossy(&attr.value);
                                    let font = if val.starts_with("major") {
                                        theme.major_font_ea.clone().or_else(|| theme.major_font.clone())
                                    } else {
                                        theme.minor_font_ea.clone().or_else(|| theme.minor_font.clone())
                                    };
                                    if let Some(f) = font {
                                        run_style.font_family_east_asia = Some(f);
                                        has_run_style = true;
                                    }
                                }
                            }
                        }
                        depth += 1;
                    }
                    _ => {
                        depth += 1;
                    }
                }
            }
            Event::Empty(e) => {
                let local = local_name(e.name().as_ref());
                if in_num_pr {
                    match local.as_str() {
                        "numId" => {
                            for attr in e.attributes().flatten() {
                                if local_name(attr.key.as_ref()) == "val" {
                                    let val = String::from_utf8_lossy(&attr.value).to_string();
                                    style.num_id = Some(val);
                                }
                            }
                        }
                        "ilvl" => {
                            for attr in e.attributes().flatten() {
                                if local_name(attr.key.as_ref()) == "val" {
                                    let val = String::from_utf8_lossy(&attr.value);
                                    style.num_ilvl = val.parse().unwrap_or(0);
                                }
                            }
                        }
                        _ => {}
                    }
                } else if in_rpr {
                    match local.as_str() {
                        "b" => {
                            run_style.bold = ct_on_off(&e);
                            run_style.has_explicit_bold = true;
                            has_run_style = true;
                        }
                        "i" => {
                            run_style.italic = ct_on_off(&e);
                            run_style.has_explicit_italic = true;
                            has_run_style = true;
                        }
                        "u" => {
                            run_style.underline = true;
                            has_run_style = true;
                            for attr in e.attributes().flatten() {
                                if local_name(attr.key.as_ref()) == "val" {
                                    let val = String::from_utf8_lossy(&attr.value).to_string();
                                    if val == "none" {
                                        run_style.underline = false;
                                    } else {
                                        run_style.underline_style = Some(val);
                                    }
                                }
                            }
                        }
                        "strike" => {
                            // S988A: honor w:val="false" (CT_OnOff three-state).
                            if std::env::var("OXI_S988A_DISABLE").is_err() {
                                run_style.strikethrough = ct_on_off(&e);
                                run_style.has_explicit_strikethrough = true;
                            } else {
                                run_style.strikethrough = true;
                            }
                            has_run_style = true;
                        }
                        "dstrike" => {
                            if std::env::var("OXI_S988A_DISABLE").is_err() {
                                let v = ct_on_off(&e);
                                run_style.strikethrough = v;
                                run_style.double_strikethrough = v;
                                run_style.has_explicit_strikethrough = true;
                                run_style.has_explicit_double_strikethrough = true;
                            } else {
                                run_style.strikethrough = true;
                                run_style.double_strikethrough = true;
                            }
                            has_run_style = true;
                        }
                        "sz" => {
                            for attr in e.attributes().flatten() {
                                if local_name(attr.key.as_ref()) == "val" {
                                    let val = String::from_utf8_lossy(&attr.value);
                                    run_style.font_size =
                                        val.parse::<f32>().ok().map(|v| v / 2.0);
                                    has_run_style = true;
                                }
                            }
                        }
                        "color" => {
                            let mut color_val = None;
                            let mut tc = None;
                            let mut tt = None;
                            let mut ts = None;
                            for attr in e.attributes().flatten() {
                                let key = local_name(attr.key.as_ref());
                                let val = String::from_utf8_lossy(&attr.value).to_string();
                                match key.as_str() {
                                    "val" => color_val = Some(val),
                                    "themeColor" => tc = Some(val),
                                    "themeTint" => tt = Some(val),
                                    "themeShade" => ts = Some(val),
                                    _ => {}
                                }
                            }
                            if let Some(ref tc_val) = tc {
                                if let Some(resolved) = theme.resolve(tc_val) {
                                    let mut hex = resolved.clone();
                                    if let Some(ref tint) = tt {
                                        if let Ok(t) = u8::from_str_radix(tint, 16) {
                                            hex = ThemeColors::apply_theme_tint(&hex, t);
                                        }
                                    }
                                    if let Some(ref shade) = ts {
                                        if let Ok(s) = u8::from_str_radix(shade, 16) {
                                            hex = ThemeColors::apply_tint_shade(&hex, -(1.0 - s as f64 / 255.0));
                                        }
                                    }
                                    run_style.color = Some(hex);
                                    has_run_style = true;
                                } else if let Some(ref cv) = color_val {
                                    if cv != "auto" { run_style.color = Some(cv.clone()); has_run_style = true; }
                                }
                            } else if let Some(ref cv) = color_val {
                                if cv != "auto" { run_style.color = Some(cv.clone()); has_run_style = true; }
                            }
                        }
                        "rFonts" => {
                            for attr in e.attributes().flatten() {
                                let key = local_name(attr.key.as_ref());
                                if key == "ascii" || key == "hAnsi" {
                                    run_style.font_family =
                                        Some(String::from_utf8_lossy(&attr.value).to_string());
                                    has_run_style = true;
                                } else if key == "asciiTheme" || key == "hAnsiTheme" {
                                    if run_style.font_family.is_none() {
                                        let val = String::from_utf8_lossy(&attr.value);
                                        let font = if val.starts_with("major") {
                                            theme.major_font.clone()
                                        } else {
                                            theme.minor_font.clone()
                                        };
                                        if let Some(f) = font {
                                            run_style.font_family = Some(f);
                                            has_run_style = true;
                                        }
                                    }
                                } else if key == "eastAsia" {
                                    run_style.font_family_east_asia =
                                        Some(String::from_utf8_lossy(&attr.value).to_string());
                                    run_style.has_explicit_east_asia = true;
                                    has_run_style = true;
                                } else if key == "eastAsiaTheme" {
                                    if run_style.font_family_east_asia.is_none() {
                                        let val = String::from_utf8_lossy(&attr.value);
                                        let font = if val.starts_with("major") {
                                            theme.major_font_ea.clone().or_else(|| theme.major_font.clone())
                                        } else {
                                            theme.minor_font_ea.clone().or_else(|| theme.minor_font.clone())
                                        };
                                        if let Some(f) = font {
                                            run_style.font_family_east_asia = Some(f);
                                            has_run_style = true;
                                        }
                                    }
                                }
                            }
                        }
                        "spacing" => {
                            for attr in e.attributes().flatten() {
                                if local_name(attr.key.as_ref()) == "val" {
                                    let val = String::from_utf8_lossy(&attr.value);
                                    run_style.character_spacing =
                                        val.parse::<f32>().ok().map(|v| v / 20.0);
                                    has_run_style = true;
                                }
                            }
                        }
                        "smallCaps" => {
                            if std::env::var("OXI_S988A_DISABLE").is_err() {
                                run_style.small_caps = ct_on_off(&e);
                                run_style.has_explicit_small_caps = true;
                            } else {
                                run_style.small_caps = true;
                            }
                            has_run_style = true;
                        }
                        "caps" => {
                            if std::env::var("OXI_S988A_DISABLE").is_err() {
                                run_style.all_caps = ct_on_off(&e);
                                run_style.has_explicit_all_caps = true;
                            } else {
                                run_style.all_caps = true;
                            }
                            has_run_style = true;
                        }
                        "shd" => {
                            for attr in e.attributes().flatten() {
                                let key = local_name(attr.key.as_ref());
                                if key == "fill" {
                                    let val = String::from_utf8_lossy(&attr.value).to_string();
                                    if val != "auto" {
                                        run_style.shading = Some(val);
                                        has_run_style = true;
                                    }
                                }
                            }
                        }
                        "bCs" => {
                            run_style.bold_cs = true;
                            has_run_style = true;
                        }
                        "iCs" => {
                            run_style.italic_cs = true;
                            has_run_style = true;
                        }
                        "szCs" => {
                            for attr in e.attributes().flatten() {
                                if local_name(attr.key.as_ref()) == "val" {
                                    let val = String::from_utf8_lossy(&attr.value);
                                    run_style.font_size_cs = val.parse::<f32>().ok().map(|v| v / 2.0);
                                    has_run_style = true;
                                }
                            }
                        }
                        "kern" => {
                            for attr in e.attributes().flatten() {
                                if local_name(attr.key.as_ref()) == "val" {
                                    let val = String::from_utf8_lossy(&attr.value);
                                    run_style.kern = val.parse::<f32>().ok().map(|v| v / 2.0);
                                    has_run_style = true;
                                }
                            }
                        }
                        // Round 29: vertAlign for character styles. The
                        // "footnote reference" character style sets
                        // <w:vertAlign w:val="superscript"/>. Without this
                        // branch, the inline rPr parser inside
                        // parse_style_definition would silently drop it,
                        // even though apply_run_property_empty handles it.
                        "vertAlign" => {
                            for attr in e.attributes().flatten() {
                                if local_name(attr.key.as_ref()) == "val" {
                                    let val = String::from_utf8_lossy(&attr.value);
                                    run_style.vertical_align = match val.as_ref() {
                                        "superscript" => Some(crate::ir::VerticalAlign::Superscript),
                                        "subscript" => Some(crate::ir::VerticalAlign::Subscript),
                                        _ => None,
                                    };
                                    if run_style.vertical_align.is_some() {
                                        has_run_style = true;
                                    }
                                }
                            }
                        }
                        _ => {}
                    }
                } else {
                    match local.as_str() {
                        "basedOn" => {
                            for attr in e.attributes().flatten() {
                                if local_name(attr.key.as_ref()) == "val" {
                                    based_on = Some(
                                        String::from_utf8_lossy(&attr.value).to_string(),
                                    );
                                }
                            }
                        }
                        "contextualSpacing" => {
                            let mut enabled = true;
                            for attr in e.attributes().flatten() {
                                if local_name(attr.key.as_ref()) == "val" {
                                    let val = String::from_utf8_lossy(&attr.value);
                                    enabled =
                                        val.as_ref() != "0" && val.as_ref() != "false";
                                }
                            }
                            style.contextual_spacing = enabled;
                        }
                        "spacing" => {
                            let mut line_val: Option<f32> = None;
                            let mut line_rule: Option<String> = None;
                            for attr in e.attributes().flatten() {
                                let key = local_name(attr.key.as_ref());
                                let val = String::from_utf8_lossy(&attr.value);
                                match key.as_str() {
                                    "before" => {
                                        style.space_before =
                                            val.parse::<f32>().ok().map(|v| v / 20.0);
                                    }
                                    "after" => {
                                        style.space_after =
                                            val.parse::<f32>().ok().map(|v| v / 20.0);
                                    }
                                    "beforeLines" => {
                                        style.before_lines = val.parse::<f32>().ok();
                                    }
                                    "afterLines" => {
                                        style.after_lines = val.parse::<f32>().ok();
                                    }
                                    // S864 retains style-level autospacing for the
                                    // table-cell parser; body inheritance stays off.
                                    "beforeAutospacing" if crate::layout::s864_part("F") => {
                                        style.before_autospacing =
                                            val == "1" || val == "true" || val == "on";
                                    }
                                    "afterAutospacing" if crate::layout::s864_part("F") => {
                                        style.after_autospacing =
                                            val == "1" || val == "true" || val == "on";
                                    }
                                    "line" => {
                                        line_val = val.parse::<f32>().ok();
                                    }
                                    "lineRule" => {
                                        line_rule = Some(val.to_string());
                                    }
                                    _ => {}
                                }
                            }
                            if let Some(lv) = line_val {
                                match line_rule.as_deref() {
                                    Some("exact") => {
                                        style.line_spacing = Some(lv / 20.0);
                                        style.line_spacing_rule = Some("exact".to_string());
                                    }
                                    Some("atLeast") => {
                                        style.line_spacing = Some(lv / 20.0);
                                        style.line_spacing_rule = Some("atLeast".to_string());
                                    }
                                    _ if lv < 0.0 => {
                                        // R7.72: negative w:line + auto → exact |val|/20
                                        style.line_spacing = Some(lv.abs() / 20.0);
                                        style.line_spacing_rule = Some("exact".to_string());
                                    }
                                    _ => {
                                        style.line_spacing = Some(lv / 240.0);
                                    }
                                }
                            }
                        }
                        "ind" => {
                            for attr in e.attributes().flatten() {
                                let key = local_name(attr.key.as_ref());
                                let val = String::from_utf8_lossy(&attr.value);
                                match key.as_str() {
                                    "left" => {
                                        style.indent_left =
                                            val.parse::<f32>().ok().map(|v| v / 20.0);
                                    }
                                    "right" => {
                                        style.indent_right =
                                            val.parse::<f32>().ok().map(|v| v / 20.0);
                                    }
                                    "firstLine" => {
                                        style.indent_first_line =
                                            val.parse::<f32>().ok().map(|v| v / 20.0);
                                    }
                                    "hanging" => {
                                        // Hanging indent: negative first-line indent
                                        style.indent_first_line =
                                            val.parse::<f32>().ok().map(|v| -(v / 20.0));
                                    }
                                    _ => {}
                                }
                            }
                        }
                        "shd" => {
                            for attr in e.attributes().flatten() {
                                if local_name(attr.key.as_ref()) == "fill" {
                                    let val = String::from_utf8_lossy(&attr.value).to_string();
                                    if val != "auto" {
                                        style.shading = Some(val);
                                    }
                                }
                            }
                        }
                        "snapToGrid" => {
                            let mut enabled = true;
                            for attr in e.attributes().flatten() {
                                if local_name(attr.key.as_ref()) == "val" {
                                    let val = String::from_utf8_lossy(&attr.value);
                                    enabled = val.as_ref() != "0" && val.as_ref() != "false";
                                }
                            }
                            style.snap_to_grid = enabled;
                        }
                        // Session 85 fix: parse autoSpaceDE/autoSpaceDN/wordWrap in
                        // STYLE definitions. Previously parse_style_definition only
                        // handled snapToGrid + widowControl among the CJK formatting
                        // group, silently losing autoSpaceDE=0 settings on styles like
                        // "ac" used by tokumei_08_01 series (a1d6/d4d126/de6e/22 docs).
                        // Result: paragraphs using <w:pStyle w:val="ac"/> retained
                        // default auto_space_de=true → 4 ASCII/CJK boundary chars ×
                        // 2.5pt = 10pt extra per line → line wrap overflow.
                        // CR9 verification: paragraph with BOTH pStyle ac AND direct
                        // pPr autoSpaceDE=0 fits 36 chars matching Word; CR6 with only
                        // pStyle ac fits 35 chars → style parser bug confirmed.
                        "autoSpaceDE" => {
                            let mut enabled = true;
                            for attr in e.attributes().flatten() {
                                if local_name(attr.key.as_ref()) == "val" {
                                    let val = String::from_utf8_lossy(&attr.value);
                                    enabled = val.as_ref() != "0" && val.as_ref() != "false";
                                }
                            }
                            style.auto_space_de = enabled;
                        }
                        "autoSpaceDN" => {
                            let mut enabled = true;
                            for attr in e.attributes().flatten() {
                                if local_name(attr.key.as_ref()) == "val" {
                                    let val = String::from_utf8_lossy(&attr.value);
                                    enabled = val.as_ref() != "0" && val.as_ref() != "false";
                                }
                            }
                            style.auto_space_dn = enabled;
                        }
                        "wordWrap" => {
                            let mut enabled = true;
                            for attr in e.attributes().flatten() {
                                if local_name(attr.key.as_ref()) == "val" {
                                    let val = String::from_utf8_lossy(&attr.value);
                                    enabled = val.as_ref() != "0" && val.as_ref() != "false";
                                }
                            }
                            style.word_wrap = enabled;
                        }
                        "pageBreakBefore" => {
                            // CT_OnOff: respect w:val="0"/"false"/"off" (S597). This
                            // is the second style-pPr handler (aiguideline's Heading
                            // styles route through here); the presence-only parse made
                            // <w:pageBreakBefore w:val="0"/> force a break → 3-page
                            // doc → 9 pages.
                            let mut enabled = true;
                            for attr in e.attributes().flatten() {
                                if local_name(attr.key.as_ref()) == "val" {
                                    let val = String::from_utf8_lossy(&attr.value);
                                    enabled = val.as_ref() != "0" && val.as_ref() != "false" && val.as_ref() != "off";
                                }
                            }
                            style.page_break_before = enabled;
                            if std::env::var("OXI_S985_DISABLE").is_err() {
                                style.has_explicit_page_break_before = true; // S985 (see other site)
                            }
                        }
                        "keepNext" => {
                            let mut enabled = true;
                            for attr in e.attributes().flatten() {
                                if local_name(attr.key.as_ref()) == "val" {
                                    let val = String::from_utf8_lossy(&attr.value);
                                    enabled = val.as_ref() != "0" && val.as_ref() != "false" && val.as_ref() != "off";
                                }
                            }
                            style.keep_next = enabled;
                            style.has_explicit_keep_next = true;
                        }
                        "keepLines" => {
                            let mut enabled = true;
                            for attr in e.attributes().flatten() {
                                if local_name(attr.key.as_ref()) == "val" {
                                    let val = String::from_utf8_lossy(&attr.value);
                                    enabled = val.as_ref() != "0" && val.as_ref() != "false" && val.as_ref() != "off";
                                }
                            }
                            style.keep_lines = enabled;
                            style.has_explicit_keep_lines = true;
                        }
                        "widowControl" => {
                            let mut enabled = true;
                            for attr in e.attributes().flatten() {
                                if local_name(attr.key.as_ref()) == "val" {
                                    let val = String::from_utf8_lossy(&attr.value);
                                    enabled = val.as_ref() != "0" && val.as_ref() != "false";
                                }
                            }
                            style.widow_control = enabled;
                            style.has_explicit_widow_control = true;
                        }
                        "jc" => {
                            for attr in e.attributes().flatten() {
                                if local_name(attr.key.as_ref()) == "val" {
                                    let val = String::from_utf8_lossy(&attr.value);
                                    alignment = Some(match val.as_ref() {
                                        "left" | "start" => Alignment::Left,
                                        "center" => Alignment::Center,
                                        "right" | "end" => Alignment::Right,
                                        "both" => Alignment::Justify,
                                        "distribute" => Alignment::Distribute,
                                        _ => Alignment::Left,
                                    });
                                }
                            }
                        }
                        "bidi" => {
                            let mut enabled = true;
                            for attr in e.attributes().flatten() {
                                if local_name(attr.key.as_ref()) == "val" {
                                    let val = String::from_utf8_lossy(&attr.value);
                                    enabled = val.as_ref() != "0" && val.as_ref() != "false";
                                }
                            }
                            style.bidi = enabled;
                        }
                        "outlineLvl" => {
                            for attr in e.attributes().flatten() {
                                if local_name(attr.key.as_ref()) == "val" {
                                    let val = String::from_utf8_lossy(&attr.value);
                                    // outlineLvl is for TOC, not layout font size
                                    style.outline_level = val.parse::<u8>().ok();
                                }
                            }
                        }
                        _ => {}
                    }
                }
            }
            Event::End(e) => {
                let local = local_name(e.name().as_ref());
                if local == "numPr" && in_num_pr {
                    in_num_pr = false;
                    depth -= 1;
                } else if local == "rPr" && in_rpr {
                    in_rpr = false;
                    depth -= 1;
                } else if local == "style" && depth == 0 {
                    break;
                } else if depth > 0 {
                    depth -= 1;
                }
            }
            Event::Eof => break,
            _ => {}
        }
    }

    if has_run_style {
        style.default_run_style = Some(run_style);
    }

    Ok((style, based_on, alignment))
}

/// Resolve table style basedOn chains.
/// Matches Word output: Table Grid (a7) basedOn Normal Table (a1)
/// → inherits borders AND cell margins from parent.
fn resolve_table_style_inheritance(styles: &mut StyleSheet) {
    let ids: Vec<String> = styles.table_styles.keys().cloned().collect();
    for id in ids {
        // style_id field stores basedOn reference (set by parse_table_style_definition)
        let based_on = styles.table_styles.get(&id)
            .and_then(|s| s.style_id.clone());
        if let Some(parent_id) = based_on {
            if let Some(parent) = styles.table_styles.get(&parent_id).cloned() {
                if let Some(child) = styles.table_styles.get_mut(&id) {
                    // Merge: child takes precedence, parent fills missing
                    if !child.border && parent.border {
                        child.border = true;
                        if child.border_color.is_none() { child.border_color = parent.border_color; }
                        if child.border_width.is_none() { child.border_width = parent.border_width; }
                        if child.border_style.is_none() { child.border_style = parent.border_style; }
                    }
                    if child.inside_horizontal_border.is_none() {
                        child.has_inside_h = parent.has_inside_h;
                        child.inside_horizontal_border = parent.inside_horizontal_border;
                    }
                    if child.inside_vertical_border.is_none() {
                        child.has_inside_v = parent.has_inside_v;
                        child.inside_vertical_border = parent.inside_vertical_border;
                    }
                    if child.default_cell_margins.is_none() {
                        child.default_cell_margins = parent.default_cell_margins;
                    }
                    if child.run_font_size.is_none() {
                        child.run_font_size = parent.run_font_size;
                    }
                }
            }
        }
    }
}

/// Parse a table style definition (type="table") from styles.xml.
/// Extracts tblBorders, tblCellMar, and tblStylePr conditional formats.
/// Returns (TableStyle, conditional_formats_map).
fn parse_table_style_definition(reader: &mut Reader<&[u8]>) -> Result<(TableStyle, std::collections::HashMap<String, TableConditionalFormat>), ParseError> {
    let mut style = TableStyle::default();
    let mut conditional_formats: std::collections::HashMap<String, TableConditionalFormat> = std::collections::HashMap::new();
    let mut depth = 0u32;
    let mut in_tbl_pr = false;
    let mut in_borders = false;
    let mut based_on: Option<String> = None;

    loop {
        match reader.read_event()? {
            Event::Start(e) => {
                let local = local_name(e.name().as_ref());
                match local.as_str() {
                    "pPr" if depth == 0 && !in_tbl_pr => {
                        // Table style base paragraph properties
                        let mut ps = crate::ir::ParagraphStyle::default();
                        let mut para_jc: Option<crate::ir::Alignment> = None;
                        let mut ppr_depth = 0u32;
                        loop {
                            match reader.read_event()? {
                                Event::Empty(pe) => {
                                    let pl = local_name(pe.name().as_ref());
                                    match pl.as_str() {
                                        "spacing" => {
                                            let mut line_val: Option<f32> = None;
                                            let mut line_rule: Option<String> = None;
                                            let mut before: Option<f32> = None;
                                            let mut after: Option<f32> = None;
                                            for attr in pe.attributes().flatten() {
                                                let key = local_name(attr.key.as_ref());
                                                let val = String::from_utf8_lossy(&attr.value);
                                                match key.as_str() {
                                                    "line" => line_val = val.parse::<f32>().ok(),
                                                    "lineRule" => line_rule = Some(val.to_string()),
                                                    "before" => before = val.parse::<f32>().ok().map(|v| v / 20.0),
                                                    "after" => after = val.parse::<f32>().ok().map(|v| v / 20.0),
                                                    // S952 (2026-07-20): a TABLE style's pPr autospacing
                                                    // applies to cell paragraphs with the S882 model
                                                    // (AUTO overrides explicit, edge-suppressed) — the
                                                    // tsas_probe derivation (001b0c6e EDU-Basic).
                                                    "beforeAutospacing" => {
                                                        ps.before_autospacing =
                                                            val == "1" || val == "true" || val == "on";
                                                    }
                                                    "afterAutospacing" => {
                                                        ps.after_autospacing =
                                                            val == "1" || val == "true" || val == "on";
                                                    }
                                                    _ => {}
                                                }
                                            }
                                            if let Some(lv) = line_val {
                                                match line_rule.as_deref() {
                                                    Some("exact") => {
                                                        ps.line_spacing = Some(lv / 20.0);
                                                        ps.line_spacing_rule = Some("exact".to_string());
                                                    }
                                                    Some("atLeast") => {
                                                        ps.line_spacing = Some(lv / 20.0);
                                                        ps.line_spacing_rule = Some("atLeast".to_string());
                                                    }
                                                    _ if lv < 0.0 => {
                                                        // R7.72: negative w:line + auto → exact |val|/20
                                                        ps.line_spacing = Some(lv.abs() / 20.0);
                                                        ps.line_spacing_rule = Some("exact".to_string());
                                                    }
                                                    _ => {
                                                        ps.line_spacing = Some(lv / 240.0);
                                                    }
                                                }
                                            }
                                            if before.is_some() { ps.space_before = before; }
                                            if after.is_some() { ps.space_after = after; }
                                        }
                                        "jc" => {
                                            for attr in pe.attributes().flatten() {
                                                if local_name(attr.key.as_ref()) == "val" {
                                                    let val = String::from_utf8_lossy(&attr.value);
                                                    para_jc = Some(match val.as_ref() {
                                                        "left" | "start" => crate::ir::Alignment::Left,
                                                        "center" => crate::ir::Alignment::Center,
                                                        "right" | "end" => crate::ir::Alignment::Right,
                                                        "both" => crate::ir::Alignment::Justify,
                                                        "distribute" => crate::ir::Alignment::Distribute,
                                                        _ => crate::ir::Alignment::Left,
                                                    });
                                                }
                                            }
                                        }
                                        "ind" => {
                                            // Table style indent: overridden by paragraph's own ind (element-level replacement)
                                        }
                                        _ => {}
                                    }
                                }
                                Event::Start(_) => { ppr_depth += 1; }
                                Event::End(pe) => {
                                    if local_name(pe.name().as_ref()) == "pPr" && ppr_depth == 0 { break; }
                                    if ppr_depth > 0 { ppr_depth -= 1; }
                                }
                                Event::Eof => break,
                                _ => {}
                            }
                        }
                        style.para_style = Some(ps);
                        style.para_alignment = para_jc;
                        continue;
                    }
                    "rPr" if depth == 0 && !in_tbl_pr => {
                        // S935: table style base run properties. w:sz here
                        // applies to every run in the table (above
                        // docDefaults, below the paragraph-style chain).
                        let mut rpr_depth = 0u32;
                        loop {
                            match reader.read_event()? {
                                Event::Empty(pe) => {
                                    if rpr_depth == 0
                                        && local_name(pe.name().as_ref()) == "sz"
                                    {
                                        for attr in pe.attributes().flatten() {
                                            if local_name(attr.key.as_ref()) == "val" {
                                                style.run_font_size = String::from_utf8_lossy(&attr.value)
                                                    .parse::<f32>()
                                                    .ok()
                                                    .map(|v| v / 2.0);
                                            }
                                        }
                                    }
                                }
                                Event::Start(_) => { rpr_depth += 1; }
                                Event::End(pe) => {
                                    if local_name(pe.name().as_ref()) == "rPr" && rpr_depth == 0 { break; }
                                    if rpr_depth > 0 { rpr_depth -= 1; }
                                }
                                Event::Eof => break,
                                _ => {}
                            }
                        }
                        continue;
                    }
                    "tblPr" if depth == 0 => { in_tbl_pr = true; }
                    "tblBorders" if in_tbl_pr => { in_borders = true; }
                    "tblCellMar" if in_tbl_pr => {
                        let mut margins = crate::ir::CellMargins { top: None, bottom: None, left: None, right: None };
                        loop {
                            match reader.read_event()? {
                                Event::Empty(me) => {
                                    let ml = local_name(me.name().as_ref());
                                    let mut w_val: Option<f32> = None;
                                    for attr in me.attributes().flatten() {
                                        if local_name(attr.key.as_ref()) == "w" {
                                            w_val = String::from_utf8_lossy(&attr.value).parse::<f32>().ok().map(|v| v / 20.0);
                                        }
                                    }
                                    match ml.as_str() {
                                        "top" => margins.top = w_val,
                                        "bottom" => margins.bottom = w_val,
                                        "left" | "start" => margins.left = w_val,
                                        "right" | "end" => margins.right = w_val,
                                        _ => {}
                                    }
                                }
                                Event::End(ee) => {
                                    if local_name(ee.name().as_ref()) == "tblCellMar" { break; }
                                }
                                Event::Eof => break,
                                _ => {}
                            }
                        }
                        style.default_cell_margins = Some(margins);
                        continue;
                    }
                    "tblStylePr" if depth == 0 => {
                        // Parse conditional table style: w:tblStylePr w:type="firstRow" etc.
                        let mut cond_type = String::new();
                        for attr in e.attributes().flatten() {
                            if local_name(attr.key.as_ref()) == "type" {
                                cond_type = String::from_utf8_lossy(&attr.value).to_string();
                            }
                        }
                        if !cond_type.is_empty() {
                            let fmt = parse_tbl_style_pr_contents(reader)?;
                            conditional_formats.insert(cond_type, fmt);
                        } else {
                            // Skip unknown tblStylePr
                            depth += 1;
                        }
                        continue;
                    }
                    _ => {}
                }
                depth += 1;
            }
            Event::Empty(e) => {
                let local = local_name(e.name().as_ref());
                match local.as_str() {
                    "basedOn" => {
                        for attr in e.attributes().flatten() {
                            if local_name(attr.key.as_ref()) == "val" {
                                based_on = Some(String::from_utf8_lossy(&attr.value).to_string());
                            }
                        }
                    }
                    "top" | "left" | "bottom" | "right" | "insideH" | "insideV" | "start" | "end"
                        if in_borders =>
                    {
                        let mut is_none = false;
                        let mut border_color = None;
                        let mut border_sz = None;
                        for attr in e.attributes().flatten() {
                            let key = local_name(attr.key.as_ref());
                            let val = String::from_utf8_lossy(&attr.value);
                            match key.as_str() {
                                "val" => { if val == "none" || val == "nil" { is_none = true; } }
                                "color" => { border_color = Some(if val == "auto" { "000000".to_string() } else { val.to_string() }); }
                                "sz" => { border_sz = val.parse::<f32>().ok().map(|v| v / 8.0); }
                                _ => {}
                            }
                        }
                        if !is_none {
                            style.border = true;
                            if local == "insideH" { style.has_inside_h = true; }
                            if local == "insideV" { style.has_inside_v = true; }
                            if style.border_color.is_none() { style.border_color = border_color.clone(); }
                            if style.border_width.is_none() { style.border_width = border_sz; }
                            style.border_style = Some("single".to_string());
                        }
                        if local == "insideH" || local == "insideV" {
                            let border = BorderDef {
                                style: if is_none { "none" } else { "single" }.to_string(),
                                width: border_sz.unwrap_or(0.5),
                                color: border_color.clone(),
                                space: 0.0,
                            };
                            if local == "insideH" {
                                style.inside_horizontal_border = Some(border);
                            } else {
                                style.inside_vertical_border = Some(border);
                            }
                        }
                    }
                    _ => {}
                }
            }
            Event::End(e) => {
                let local = local_name(e.name().as_ref());
                if local == "tblBorders" { in_borders = false; }
                if local == "tblPr" { in_tbl_pr = false; }
                if local == "style" && depth <= 1 { break; }
                if depth > 0 { depth -= 1; }
            }
            Event::Eof => break,
            _ => {}
        }
    }

    style.style_id = based_on;
    Ok((style, conditional_formats))
}

/// Parse the contents of a w:tblStylePr element.
/// Extracts tcPr (shading, borders), pPr (alignment), rPr (bold, color).
fn parse_tbl_style_pr_contents(reader: &mut Reader<&[u8]>) -> Result<TableConditionalFormat, ParseError> {
    let mut fmt = TableConditionalFormat::default();
    let mut depth = 0u32;
    let mut in_tc_pr = false;
    let mut in_tc_borders = false;
    let mut borders = CellBorders {
        top: None, bottom: None, left: None, right: None,
    };
    let mut has_borders = false;

    loop {
        match reader.read_event()? {
            Event::Start(e) => {
                let local = local_name(e.name().as_ref());
                match local.as_str() {
                    "tcPr" if depth == 0 => { in_tc_pr = true; }
                    "tcBorders" if in_tc_pr => { in_tc_borders = true; }
                    _ => {}
                }
                depth += 1;
            }
            Event::Empty(e) => {
                let local = local_name(e.name().as_ref());
                if in_tc_borders {
                    match local.as_str() {
                        "top" | "bottom" | "left" | "right" | "start" | "end" => {
                            let bdef = parse_border_def_from_attrs(&e);
                            if bdef.is_some() { has_borders = true; }
                            match local.as_str() {
                                "top" => borders.top = bdef,
                                "bottom" => borders.bottom = bdef,
                                "left" | "start" => borders.left = bdef,
                                "right" | "end" => borders.right = bdef,
                                _ => {}
                            }
                        }
                        _ => {}
                    }
                } else if in_tc_pr {
                    if local == "shd" {
                        for attr in e.attributes().flatten() {
                            let key = local_name(attr.key.as_ref());
                            if key == "fill" {
                                let val = String::from_utf8_lossy(&attr.value).to_string();
                                if val != "auto" {
                                    fmt.shading = Some(val);
                                }
                            }
                        }
                    }
                } else {
                    // Elements inside tblStylePr (at any depth, including rPr)
                    match local.as_str() {
                        "b" | "bCs" => {
                            // <w:b/> means bold=true, <w:b w:val="0"/> means bold=false
                            let mut val = true;
                            for attr in e.attributes().flatten() {
                                if local_name(attr.key.as_ref()) == "val" {
                                    let v = String::from_utf8_lossy(&attr.value);
                                    if v == "0" || v == "false" { val = false; }
                                }
                            }
                            fmt.bold = Some(val);
                        }
                        "color" => {
                            for attr in e.attributes().flatten() {
                                if local_name(attr.key.as_ref()) == "val" {
                                    let v = String::from_utf8_lossy(&attr.value).to_string();
                                    if v != "auto" { fmt.color = Some(v); }
                                }
                            }
                        }
                        "jc" => {
                            for attr in e.attributes().flatten() {
                                if local_name(attr.key.as_ref()) == "val" {
                                    let val = String::from_utf8_lossy(&attr.value);
                                    fmt.alignment = Some(match val.as_ref() {
                                        "left" | "start" => Alignment::Left,
                                        "center" => Alignment::Center,
                                        "right" | "end" => Alignment::Right,
                                        "both" => Alignment::Justify,
                                        "distribute" => Alignment::Distribute,
                                        _ => Alignment::Left,
                                    });
                                }
                            }
                        }
                        _ => {}
                    }
                }
            }
            Event::End(e) => {
                let local = local_name(e.name().as_ref());
                if local == "tcBorders" { in_tc_borders = false; }
                if local == "tcPr" { in_tc_pr = false; }
                if local == "tblStylePr" && depth <= 1 { break; }
                if depth > 0 { depth -= 1; }
            }
            Event::Eof => break,
            _ => {}
        }
    }

    if has_borders {
        fmt.borders = Some(borders);
    }
    Ok(fmt)
}

/// Parse a single border element's attributes into a BorderDef.
fn parse_border_def_from_attrs(e: &quick_xml::events::BytesStart) -> Option<BorderDef> {
    let mut is_none = false;
    let mut color = None;
    let mut width = None;
    let mut style = None;
    for attr in e.attributes().flatten() {
        let key = local_name(attr.key.as_ref());
        let val = String::from_utf8_lossy(&attr.value);
        match key.as_str() {
            "val" => {
                if val == "none" || val == "nil" { is_none = true; }
                style = Some(val.to_string());
            }
            "color" => {
                color = Some(if val == "auto" { "000000".to_string() } else { val.to_string() });
            }
            "sz" => {
                width = val.parse::<f32>().ok().map(|v| v / 8.0);
            }
            _ => {}
        }
    }
    if is_none { return None; }
    Some(BorderDef {
        color,
        width: width.unwrap_or(0.5),
        style: style.unwrap_or_else(|| "single".to_string()),
        space: 0.0,
    })
}

fn local_name(name: &[u8]) -> String {
    let s = std::str::from_utf8(name).unwrap_or("");
    match s.rfind(':') {
        Some(pos) => s[pos + 1..].to_string(),
        None => s.to_string(),
    }
}

#[cfg(test)]
mod s985_pbb_inheritance {
    // S985 (Task legal__001410a8): pageBreakBefore is CT_OnOff three-state in the
    // style chain — a derived style's explicit value (incl. val="0") beats a
    // basedOn parent, and only an unspecified child inherits.
    use crate::ir::ParagraphStyle;

    fn ps(pbb: bool, explicit: bool) -> ParagraphStyle {
        ParagraphStyle {
            page_break_before: pbb,
            has_explicit_page_break_before: explicit,
            ..Default::default()
        }
    }

    #[test]
    fn parent_true_child_unspecified_inherits_true() {
        let mut child = ps(false, false);
        super::merge_para_style(&mut child, &ps(true, true));
        assert!(child.page_break_before, "unspecified child inherits parent ON");
        assert!(child.has_explicit_page_break_before, "explicitness propagates");
    }

    #[test]
    fn parent_true_child_explicit_false_stays_false() {
        let mut child = ps(false, true);
        super::merge_para_style(&mut child, &ps(true, true));
        assert!(
            !child.page_break_before,
            "an explicit child val=0 must beat the basedOn parent's ON"
        );
    }

    #[test]
    fn parent_false_child_explicit_true_stays_true() {
        let mut child = ps(true, true);
        super::merge_para_style(&mut child, &ps(false, true));
        assert!(child.page_break_before, "explicit child true survives parent OFF");
    }
}

