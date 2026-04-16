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
    } else {
        // "majorHAnsi" / "minorHAnsi" / "majorBidi" / "minorBidi" → Latin font
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

        if let Some(parent) = parent_para {
            let def = styles.styles.get_mut(id).unwrap();
            // Merge: child overrides parent. Only fill in None/default fields from parent.
            merge_para_style(&mut def.paragraph, &parent);
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
    if child.indent_left.is_none() {
        child.indent_left = parent.indent_left;
    }
    if child.indent_right.is_none() {
        child.indent_right = parent.indent_right;
    }
    if child.indent_first_line.is_none() {
        child.indent_first_line = parent.indent_first_line;
    }
    if child.num_id.is_none() {
        child.num_id = parent.num_id.clone();
        if child.num_id.is_some() {
            child.num_ilvl = parent.num_ilvl;
        }
    }
    // Inherit bool fields: true from parent if child doesn't set
    if !child.keep_next && parent.keep_next {
        child.keep_next = true;
    }
    if !child.keep_lines && parent.keep_lines {
        child.keep_lines = true;
    }
    if !child.has_explicit_widow_control {
        child.widow_control = parent.widow_control;
    }
    if !child.page_break_before && parent.page_break_before {
        child.page_break_before = true;
    }
    if !child.contextual_spacing && parent.contextual_spacing {
        child.contextual_spacing = true;
    }
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
    if child.tab_stops.is_empty() {
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
    // Boolean properties: inherit from parent if child doesn't explicitly set them
    if !child.bold && parent.bold {
        child.bold = true;
    }
    if !child.italic && parent.italic {
        child.italic = true;
    }
    if !child.underline && parent.underline {
        child.underline = true;
    }
    if !child.strikethrough && parent.strikethrough {
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

/// Apply a single empty run property element to RunStyle
fn apply_run_property_empty(e: &quick_xml::events::BytesStart, rs: &mut RunStyle, theme: &ThemeColors) {
    let local = local_name(e.name().as_ref());
    match local.as_str() {
        "b" => rs.bold = true,
        "i" => rs.italic = true,
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
        "strike" => rs.strikethrough = true,
        "dstrike" => {
            rs.strikethrough = true;
            rs.double_strikethrough = true;
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
                            let tint_val = t as f64 / 255.0;
                            hex = ThemeColors::apply_tint_shade(&hex, tint_val);
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
        "spacing" => {
            for attr in e.attributes().flatten() {
                if local_name(attr.key.as_ref()) == "val" {
                    let val = String::from_utf8_lossy(&attr.value);
                    rs.character_spacing = val.parse::<f32>().ok().map(|v| v / 20.0);
                }
            }
        }
        "smallCaps" => rs.small_caps = true,
        "caps" => rs.all_caps = true,
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
        "vanish" | "webHidden" => rs.vanish = true,
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
        "pageBreakBefore" => {
            style.page_break_before = true;
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
        "keepNext" => style.keep_next = true,
        "keepLines" => style.keep_lines = true,
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
                            run_style.bold = true;
                            has_run_style = true;
                        }
                        "i" => {
                            run_style.italic = true;
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
                            run_style.strikethrough = true;
                            has_run_style = true;
                        }
                        "dstrike" => {
                            run_style.strikethrough = true;
                            run_style.double_strikethrough = true;
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
                                            hex = ThemeColors::apply_tint_shade(&hex, t as f64 / 255.0);
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
                            run_style.small_caps = true;
                            has_run_style = true;
                        }
                        "caps" => {
                            run_style.all_caps = true;
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
                        "pageBreakBefore" => {
                            style.page_break_before = true;
                        }
                        "keepNext" => style.keep_next = true,
                        "keepLines" => style.keep_lines = true,
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
                    if child.default_cell_margins.is_none() {
                        child.default_cell_margins = parent.default_cell_margins;
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
                            let local_border = local_name(e.name().as_ref());
                            if local_border == "insideH" {
                                style.has_inside_h = true;
                            }
                            if style.border_color.is_none() { style.border_color = border_color; }
                            if style.border_width.is_none() { style.border_width = border_sz; }
                            style.border_style = Some("single".to_string());
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

