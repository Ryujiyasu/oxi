use quick_xml::events::Event;
use quick_xml::reader::Reader;

use super::ParseError;
use crate::ir::{ParagraphStyle, RunStyle, StyleDefinition, StyleSheet};

pub fn parse_styles(xml: &str) -> Result<StyleSheet, ParseError> {
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
                        let run_style = parse_run_properties_block(&mut reader)?;
                        styles.doc_default_run_style = Some(run_style);
                    }
                    "pPr" if in_ppr_default => {
                        let para_style = parse_para_properties_block(&mut reader)?;
                        styles.doc_default_para_style = Some(para_style);
                    }
                    "style" => {
                        let mut style_id = None;
                        let mut style_type = None;
                        for attr in e.attributes().flatten() {
                            let key = local_name(attr.key.as_ref());
                            let val = String::from_utf8_lossy(&attr.value).to_string();
                            match key.as_str() {
                                "styleId" => style_id = Some(val),
                                "type" => style_type = Some(val),
                                _ => {}
                            }
                        }
                        if let (Some(id), Some(typ)) = (style_id, style_type) {
                            if typ == "paragraph" || typ == "character" {
                                let (pstyle, based_on) = parse_style_definition(&mut reader)?;
                                styles.styles.insert(
                                    id.clone(),
                                    StyleDefinition {
                                        style_id: id,
                                        based_on,
                                        paragraph: pstyle,
                                        resolved: false,
                                    },
                                );
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
    if child.indent_left.is_none() {
        child.indent_left = parent.indent_left;
    }
    if child.indent_right.is_none() {
        child.indent_right = parent.indent_right;
    }
    if child.indent_first_line.is_none() {
        child.indent_first_line = parent.indent_first_line;
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
fn merge_run_style(child: &mut RunStyle, parent: &RunStyle) {
    if child.font_family.is_none() {
        child.font_family = parent.font_family.clone();
    }
    if child.font_family_east_asia.is_none() {
        child.font_family_east_asia = parent.font_family_east_asia.clone();
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
}

/// Parse a run properties block (<w:rPr>...</w:rPr>) and return RunStyle
fn parse_run_properties_block(reader: &mut Reader<&[u8]>) -> Result<RunStyle, ParseError> {
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
                        } else if key == "eastAsia" {
                            rs.font_family_east_asia =
                                Some(String::from_utf8_lossy(&attr.value).to_string());
                        }
                    }
                }
                depth += 1;
            }
            Event::Empty(e) => {
                apply_run_property_empty(&e, &mut rs);
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

/// Parse a paragraph properties block (<w:pPr>...</w:pPr>) and return ParagraphStyle
fn parse_para_properties_block(reader: &mut Reader<&[u8]>) -> Result<ParagraphStyle, ParseError> {
    let mut style = ParagraphStyle::default();
    let mut depth = 1;

    loop {
        match reader.read_event()? {
            Event::Start(_) => {
                depth += 1;
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
fn apply_run_property_empty(e: &quick_xml::events::BytesStart, rs: &mut RunStyle) {
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
            for attr in e.attributes().flatten() {
                if local_name(attr.key.as_ref()) == "val" {
                    let val = String::from_utf8_lossy(&attr.value);
                    if val != "auto" {
                        rs.color = Some(val.to_string());
                    }
                }
            }
        }
        "rFonts" => {
            for attr in e.attributes().flatten() {
                let key = local_name(attr.key.as_ref());
                if key == "ascii" || key == "hAnsi" {
                    rs.font_family =
                        Some(String::from_utf8_lossy(&attr.value).to_string());
                } else if key == "eastAsia" {
                    rs.font_family_east_asia =
                        Some(String::from_utf8_lossy(&attr.value).to_string());
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
) -> Result<(ParagraphStyle, Option<String>), ParseError> {
    let mut style = ParagraphStyle::default();
    let mut run_style = RunStyle::default();
    let mut has_run_style = false;
    let mut based_on: Option<String> = None;
    let mut depth = 0;
    let mut in_rpr = false;

    loop {
        match reader.read_event()? {
            Event::Start(e) => {
                let local = local_name(e.name().as_ref());
                match local.as_str() {
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
                            } else if key == "eastAsia" {
                                run_style.font_family_east_asia =
                                    Some(String::from_utf8_lossy(&attr.value).to_string());
                                has_run_style = true;
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
                if in_rpr {
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
                            for attr in e.attributes().flatten() {
                                if local_name(attr.key.as_ref()) == "val" {
                                    let val = String::from_utf8_lossy(&attr.value);
                                    if val != "auto" {
                                        run_style.color = Some(val.to_string());
                                        has_run_style = true;
                                    }
                                }
                            }
                        }
                        "rFonts" => {
                            for attr in e.attributes().flatten() {
                                let key = local_name(attr.key.as_ref());
                                if key == "ascii" || key == "hAnsi" {
                                    run_style.font_family =
                                        Some(String::from_utf8_lossy(&attr.value).to_string());
                                    has_run_style = true;
                                } else if key == "eastAsia" {
                                    run_style.font_family_east_asia =
                                        Some(String::from_utf8_lossy(&attr.value).to_string());
                                    has_run_style = true;
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
            }
            Event::End(e) => {
                let local = local_name(e.name().as_ref());
                if local == "rPr" && in_rpr {
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

    Ok((style, based_on))
}

fn local_name(name: &[u8]) -> String {
    let s = std::str::from_utf8(name).unwrap_or("");
    match s.rfind(':') {
        Some(pos) => s[pos + 1..].to_string(),
        None => s.to_string(),
    }
}
