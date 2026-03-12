use quick_xml::events::Event;
use quick_xml::reader::Reader;

use super::ParseError;
use crate::ir::{ParagraphStyle, StyleSheet};

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
    let mut depth = 0;

    loop {
        match reader.read_event()? {
            Event::Start(_e) => {
                depth += 1;
            }
            Event::Empty(e) => {
                let local = local_name(e.name().as_ref());
                if local == "spacing" {
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
                if local == "style" && depth == 0 {
                    break;
                }
                if depth > 0 {
                    depth -= 1;
                }
            }
            Event::Eof => break,
            _ => {}
        }
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
