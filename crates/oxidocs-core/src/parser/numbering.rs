use std::collections::HashMap;

use quick_xml::events::Event;
use quick_xml::reader::Reader;

use super::ParseError;

/// A single numbering level definition
#[derive(Debug, Clone)]
pub struct NumberingLevel {
    /// Level index (0-8)
    pub ilvl: u8,
    /// Number format: "bullet", "decimal", "lowerLetter", "upperLetter", "lowerRoman", "upperRoman", etc.
    pub num_fmt: String,
    /// Level text pattern, e.g., "%1." or a literal bullet character
    pub lvl_text: String,
    /// Left indent in points (from w:ind w:left, converted from twips)
    pub indent_left: Option<f32>,
    /// Hanging indent in points (from w:ind w:hanging, converted from twips)
    pub indent_hanging: Option<f32>,
}

/// An abstract numbering definition containing levels
#[derive(Debug, Clone)]
pub struct AbstractNum {
    pub abstract_num_id: String,
    pub levels: HashMap<u8, NumberingLevel>,
}

/// Parsed numbering definitions from word/numbering.xml
#[derive(Debug, Clone, Default)]
pub struct NumberingDefinitions {
    /// abstractNumId -> AbstractNum
    pub abstract_nums: HashMap<String, AbstractNum>,
    /// numId -> abstractNumId
    pub num_map: HashMap<String, String>,
}

impl NumberingDefinitions {
    /// Resolve the marker text and hanging indent for a given numId and ilvl.
    /// For numbered lists, increments the counter.
    pub fn resolve_marker(
        &self,
        num_id: &str,
        ilvl: u8,
        counters: &mut HashMap<(String, u8), u32>,
    ) -> (String, Option<f32>) {
        let abstract_num_id = match self.num_map.get(num_id) {
            Some(id) => id,
            None => return ("\u{2022}".to_string(), Some(18.0)), // fallback bullet
        };

        let abstract_num = match self.abstract_nums.get(abstract_num_id) {
            Some(an) => an,
            None => return ("\u{2022}".to_string(), Some(18.0)),
        };

        let level = match abstract_num.levels.get(&ilvl) {
            Some(l) => l,
            None => return ("\u{2022}".to_string(), Some(18.0)),
        };

        let hanging = level.indent_hanging;

        if level.num_fmt == "bullet" {
            // For bullets, use the lvlText directly
            let marker = if level.lvl_text.is_empty() {
                "\u{2022}".to_string()
            } else {
                level.lvl_text.clone()
            };
            return (marker, hanging);
        }

        // Numbered list: increment counter
        let key = (num_id.to_string(), ilvl);
        let count = counters.entry(key).or_insert(0);
        *count += 1;
        let current = *count;

        // Format the number based on num_fmt
        let formatted_num = format_number(current, &level.num_fmt);

        // Replace %1, %2, etc. in lvl_text with the formatted number
        // For simplicity, we replace %N where N = ilvl+1 with the current level's number
        let marker = if level.lvl_text.is_empty() {
            format!("{}.", formatted_num)
        } else {
            let placeholder = format!("%{}", ilvl + 1);
            level.lvl_text.replace(&placeholder, &formatted_num)
        };

        (marker, hanging)
    }

    /// Get the left indent for a given numId and ilvl
    pub fn get_level_indent(&self, num_id: &str, ilvl: u8) -> Option<f32> {
        let abstract_num_id = self.num_map.get(num_id)?;
        let abstract_num = self.abstract_nums.get(abstract_num_id)?;
        let level = abstract_num.levels.get(&ilvl)?;
        level.indent_left
    }
}

fn format_number(n: u32, fmt: &str) -> String {
    match fmt {
        "decimal" => n.to_string(),
        "lowerLetter" => {
            if n >= 1 && n <= 26 {
                ((b'a' + (n - 1) as u8) as char).to_string()
            } else {
                n.to_string()
            }
        }
        "upperLetter" => {
            if n >= 1 && n <= 26 {
                ((b'A' + (n - 1) as u8) as char).to_string()
            } else {
                n.to_string()
            }
        }
        "lowerRoman" => to_roman(n, false),
        "upperRoman" => to_roman(n, true),
        _ => n.to_string(), // fallback to decimal
    }
}

fn to_roman(mut n: u32, upper: bool) -> String {
    let values = [
        (1000, "m"),
        (900, "cm"),
        (500, "d"),
        (400, "cd"),
        (100, "c"),
        (90, "xc"),
        (50, "l"),
        (40, "xl"),
        (10, "x"),
        (9, "ix"),
        (5, "v"),
        (4, "iv"),
        (1, "i"),
    ];
    let mut result = String::new();
    for &(val, sym) in &values {
        while n >= val {
            result.push_str(sym);
            n -= val;
        }
    }
    if upper {
        result.to_uppercase()
    } else {
        result
    }
}

/// Parse word/numbering.xml
pub fn parse_numbering(xml: &str) -> Result<NumberingDefinitions, ParseError> {
    let mut reader = Reader::from_str(xml);
    let mut defs = NumberingDefinitions::default();

    loop {
        match reader.read_event()? {
            Event::Start(e) => {
                let local = local_name(e.name().as_ref());
                match local.as_str() {
                    "abstractNum" => {
                        let mut abstract_num_id = String::new();
                        for attr in e.attributes().flatten() {
                            let key = local_name(attr.key.as_ref());
                            if key == "abstractNumId" {
                                abstract_num_id =
                                    String::from_utf8_lossy(&attr.value).to_string();
                            }
                        }
                        let abs_num = parse_abstract_num(&mut reader, &abstract_num_id)?;
                        defs.abstract_nums
                            .insert(abstract_num_id, abs_num);
                    }
                    "num" => {
                        let mut num_id = String::new();
                        for attr in e.attributes().flatten() {
                            let key = local_name(attr.key.as_ref());
                            if key == "numId" {
                                num_id = String::from_utf8_lossy(&attr.value).to_string();
                            }
                        }
                        let abstract_num_id = parse_num_element(&mut reader)?;
                        defs.num_map.insert(num_id, abstract_num_id);
                    }
                    _ => {}
                }
            }
            Event::Eof => break,
            _ => {}
        }
    }

    Ok(defs)
}

/// Parse a w:abstractNum element
fn parse_abstract_num(
    reader: &mut Reader<&[u8]>,
    abstract_num_id: &str,
) -> Result<AbstractNum, ParseError> {
    let mut levels = HashMap::new();
    let mut depth = 0;

    loop {
        match reader.read_event()? {
            Event::Start(e) => {
                let local = local_name(e.name().as_ref());
                if local == "lvl" && depth == 0 {
                    let mut ilvl: u8 = 0;
                    for attr in e.attributes().flatten() {
                        let key = local_name(attr.key.as_ref());
                        if key == "ilvl" {
                            let val = String::from_utf8_lossy(&attr.value);
                            ilvl = val.parse().unwrap_or(0);
                        }
                    }
                    let level = parse_numbering_level(reader, ilvl)?;
                    levels.insert(ilvl, level);
                } else {
                    depth += 1;
                }
            }
            Event::End(e) => {
                let local = local_name(e.name().as_ref());
                if local == "abstractNum" && depth == 0 {
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

    Ok(AbstractNum {
        abstract_num_id: abstract_num_id.to_string(),
        levels,
    })
}

/// Parse a single w:lvl element
fn parse_numbering_level(
    reader: &mut Reader<&[u8]>,
    ilvl: u8,
) -> Result<NumberingLevel, ParseError> {
    let mut num_fmt = String::new();
    let mut lvl_text = String::new();
    let mut indent_left = None;
    let mut indent_hanging = None;
    let mut depth = 0;

    loop {
        match reader.read_event()? {
            Event::Start(e) => {
                let local = local_name(e.name().as_ref());
                if local == "pPr" && depth == 0 {
                    // Parse level's paragraph properties for indentation
                    depth += 1;
                } else {
                    depth += 1;
                }
            }
            Event::Empty(e) => {
                let local = local_name(e.name().as_ref());
                match local.as_str() {
                    "numFmt" => {
                        for attr in e.attributes().flatten() {
                            if local_name(attr.key.as_ref()) == "val" {
                                num_fmt =
                                    String::from_utf8_lossy(&attr.value).to_string();
                            }
                        }
                    }
                    "lvlText" => {
                        for attr in e.attributes().flatten() {
                            if local_name(attr.key.as_ref()) == "val" {
                                let raw = String::from_utf8_lossy(&attr.value).to_string();
                                lvl_text = unescape_xml_entities(&raw);
                            }
                        }
                    }
                    "ind" => {
                        for attr in e.attributes().flatten() {
                            let key = local_name(attr.key.as_ref());
                            let val = String::from_utf8_lossy(&attr.value);
                            match key.as_str() {
                                "left" => {
                                    indent_left =
                                        val.parse::<f32>().ok().map(|v| v / 20.0);
                                }
                                "hanging" => {
                                    indent_hanging =
                                        val.parse::<f32>().ok().map(|v| v / 20.0);
                                }
                                _ => {}
                            }
                        }
                    }
                    _ => {}
                }
            }
            Event::End(e) => {
                let local = local_name(e.name().as_ref());
                if local == "lvl" && depth == 0 {
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

    Ok(NumberingLevel {
        ilvl,
        num_fmt,
        lvl_text,
        indent_left,
        indent_hanging,
    })
}

/// Parse a w:num element to get its abstractNumId
fn parse_num_element(reader: &mut Reader<&[u8]>) -> Result<String, ParseError> {
    let mut abstract_num_id = String::new();
    let mut depth = 0;

    loop {
        match reader.read_event()? {
            Event::Start(_) => {
                depth += 1;
            }
            Event::Empty(e) => {
                let local = local_name(e.name().as_ref());
                if local == "abstractNumId" {
                    for attr in e.attributes().flatten() {
                        if local_name(attr.key.as_ref()) == "val" {
                            abstract_num_id =
                                String::from_utf8_lossy(&attr.value).to_string();
                        }
                    }
                }
            }
            Event::End(e) => {
                let local = local_name(e.name().as_ref());
                if local == "num" && depth == 0 {
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

    Ok(abstract_num_id)
}

/// Unescape XML character references like &#x2022; and &#8226; and standard entities
fn unescape_xml_entities(s: &str) -> String {
    let mut result = String::with_capacity(s.len());
    let mut chars = s.chars().peekable();
    while let Some(ch) = chars.next() {
        if ch == '&' {
            let mut entity = String::new();
            for c in chars.by_ref() {
                if c == ';' {
                    break;
                }
                entity.push(c);
            }
            if entity.starts_with("#x") || entity.starts_with("#X") {
                if let Ok(code) = u32::from_str_radix(&entity[2..], 16) {
                    if let Some(c) = char::from_u32(code) {
                        result.push(c);
                        continue;
                    }
                }
            } else if entity.starts_with('#') {
                if let Ok(code) = entity[1..].parse::<u32>() {
                    if let Some(c) = char::from_u32(code) {
                        result.push(c);
                        continue;
                    }
                }
            } else {
                match entity.as_str() {
                    "amp" => { result.push('&'); continue; }
                    "lt" => { result.push('<'); continue; }
                    "gt" => { result.push('>'); continue; }
                    "quot" => { result.push('"'); continue; }
                    "apos" => { result.push('\''); continue; }
                    _ => {}
                }
            }
            // Fallback: put original back
            result.push('&');
            result.push_str(&entity);
            result.push(';');
        } else {
            result.push(ch);
        }
    }
    result
}

fn local_name(name: &[u8]) -> String {
    let s = std::str::from_utf8(name).unwrap_or("");
    match s.rfind(':') {
        Some(pos) => s[pos + 1..].to_string(),
        None => s.to_string(),
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_parse_numbering_xml() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:abstractNum w:abstractNumId="0">
    <w:lvl w:ilvl="0">
      <w:numFmt w:val="bullet"/>
      <w:lvlText w:val="&#x2022;"/>
      <w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr>
    </w:lvl>
    <w:lvl w:ilvl="1">
      <w:numFmt w:val="bullet"/>
      <w:lvlText w:val="&#x25CB;"/>
      <w:pPr><w:ind w:left="1440" w:hanging="360"/></w:pPr>
    </w:lvl>
  </w:abstractNum>
  <w:abstractNum w:abstractNumId="1">
    <w:lvl w:ilvl="0">
      <w:numFmt w:val="decimal"/>
      <w:lvlText w:val="%1."/>
      <w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr>
    </w:lvl>
  </w:abstractNum>
  <w:num w:numId="1">
    <w:abstractNumId w:val="0"/>
  </w:num>
  <w:num w:numId="2">
    <w:abstractNumId w:val="1"/>
  </w:num>
</w:numbering>"#;

        let defs = parse_numbering(xml).unwrap();

        // Check abstract nums
        assert_eq!(defs.abstract_nums.len(), 2);
        assert!(defs.abstract_nums.contains_key("0"));
        assert!(defs.abstract_nums.contains_key("1"));

        // Check bullet level
        let abs0 = &defs.abstract_nums["0"];
        let lvl0 = &abs0.levels[&0];
        assert_eq!(lvl0.num_fmt, "bullet");
        assert_eq!(lvl0.lvl_text, "\u{2022}");
        assert_eq!(lvl0.indent_left, Some(36.0));   // 720 / 20
        assert_eq!(lvl0.indent_hanging, Some(18.0)); // 360 / 20

        // Check decimal level
        let abs1 = &defs.abstract_nums["1"];
        let lvl0 = &abs1.levels[&0];
        assert_eq!(lvl0.num_fmt, "decimal");
        assert_eq!(lvl0.lvl_text, "%1.");

        // Check num map
        assert_eq!(defs.num_map["1"], "0");
        assert_eq!(defs.num_map["2"], "1");
    }

    #[test]
    fn test_resolve_bullet_marker() {
        let xml = r#"<?xml version="1.0"?>
<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:abstractNum w:abstractNumId="0">
    <w:lvl w:ilvl="0">
      <w:numFmt w:val="bullet"/>
      <w:lvlText w:val="&#x2022;"/>
      <w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr>
    </w:lvl>
  </w:abstractNum>
  <w:num w:numId="1">
    <w:abstractNumId w:val="0"/>
  </w:num>
</w:numbering>"#;

        let defs = parse_numbering(xml).unwrap();
        let mut counters = HashMap::new();

        let (marker, indent) = defs.resolve_marker("1", 0, &mut counters);
        assert_eq!(marker, "\u{2022}");
        assert_eq!(indent, Some(18.0));

        // Bullet doesn't increment counter
        let (marker2, _) = defs.resolve_marker("1", 0, &mut counters);
        assert_eq!(marker2, "\u{2022}");
    }

    #[test]
    fn test_resolve_decimal_marker() {
        let xml = r#"<?xml version="1.0"?>
<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:abstractNum w:abstractNumId="0">
    <w:lvl w:ilvl="0">
      <w:numFmt w:val="decimal"/>
      <w:lvlText w:val="%1."/>
      <w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr>
    </w:lvl>
  </w:abstractNum>
  <w:num w:numId="1">
    <w:abstractNumId w:val="0"/>
  </w:num>
</w:numbering>"#;

        let defs = parse_numbering(xml).unwrap();
        let mut counters = HashMap::new();

        let (m1, _) = defs.resolve_marker("1", 0, &mut counters);
        assert_eq!(m1, "1.");

        let (m2, _) = defs.resolve_marker("1", 0, &mut counters);
        assert_eq!(m2, "2.");

        let (m3, _) = defs.resolve_marker("1", 0, &mut counters);
        assert_eq!(m3, "3.");
    }

    #[test]
    fn test_resolve_lower_letter_marker() {
        let xml = r#"<?xml version="1.0"?>
<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:abstractNum w:abstractNumId="0">
    <w:lvl w:ilvl="0">
      <w:numFmt w:val="lowerLetter"/>
      <w:lvlText w:val="%1)"/>
      <w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr>
    </w:lvl>
  </w:abstractNum>
  <w:num w:numId="1">
    <w:abstractNumId w:val="0"/>
  </w:num>
</w:numbering>"#;

        let defs = parse_numbering(xml).unwrap();
        let mut counters = HashMap::new();

        let (m1, _) = defs.resolve_marker("1", 0, &mut counters);
        assert_eq!(m1, "a)");

        let (m2, _) = defs.resolve_marker("1", 0, &mut counters);
        assert_eq!(m2, "b)");
    }

    #[test]
    fn test_roman_numerals() {
        assert_eq!(to_roman(1, false), "i");
        assert_eq!(to_roman(4, false), "iv");
        assert_eq!(to_roman(9, false), "ix");
        assert_eq!(to_roman(14, false), "xiv");
        assert_eq!(to_roman(3, true), "III");
    }

    #[test]
    fn test_unknown_num_id_fallback() {
        let defs = NumberingDefinitions::default();
        let mut counters = HashMap::new();

        let (marker, indent) = defs.resolve_marker("999", 0, &mut counters);
        assert_eq!(marker, "\u{2022}");
        assert_eq!(indent, Some(18.0));
    }
}
