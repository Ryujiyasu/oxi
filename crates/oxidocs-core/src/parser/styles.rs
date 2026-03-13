use quick_xml::events::Event;
use quick_xml::reader::Reader;

use super::ParseError;
use crate::ir::{ParagraphStyle, RunStyle, StyleSheet};

pub fn parse_styles(xml: &str) -> Result<StyleSheet, ParseError> {
    let mut reader = Reader::from_str(xml);
    let mut styles = StyleSheet::default();

    loop {
        match reader.read_event()? {
            Event::Start(e) => {
                let local = local_name(e.name().as_ref());
                if local == "style" {
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
                        if typ == "paragraph" {
                            let pstyle = parse_style_definition(&mut reader)?;
                            styles.styles.insert(id, pstyle);
                        }
                    }
                }
            }
            Event::Eof => break,
            _ => {}
        }
    }

    Ok(styles)
}

fn parse_style_definition(
    reader: &mut Reader<&[u8]>,
) -> Result<ParagraphStyle, ParseError> {
    let mut style = ParagraphStyle::default();
    let mut run_style = RunStyle::default();
    let mut has_run_style = false;
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
                        _ => {}
                    }
                } else if local == "contextualSpacing" {
                    let mut enabled = true;
                    for attr in e.attributes().flatten() {
                        if local_name(attr.key.as_ref()) == "val" {
                            let val = String::from_utf8_lossy(&attr.value);
                            enabled = val.as_ref() != "0" && val.as_ref() != "false";
                        }
                    }
                    style.contextual_spacing = enabled;
                } else if local == "spacing" {
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
                                style.line_spacing = val.parse::<f32>().ok().map(|v| v / 240.0);
                            }
                            _ => {}
                        }
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

    Ok(style)
}

fn local_name(name: &[u8]) -> String {
    let s = std::str::from_utf8(name).unwrap_or("");
    match s.rfind(':') {
        Some(pos) => s[pos + 1..].to_string(),
        None => s.to_string(),
    }
}
