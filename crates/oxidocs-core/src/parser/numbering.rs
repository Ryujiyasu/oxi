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
    /// Suffix after number: "tab" (default), "space", or "nothing"
    pub suff: String,
    /// Tab stop position in points (from w:tabs/w:tab w:pos, converted from twips)
    pub tab_stop: Option<f32>,
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
    /// numId -> (ilvl -> start override value)
    pub level_overrides: HashMap<String, HashMap<u8, u32>>,
}

/// Result of resolving a list marker
pub struct ResolvedMarker {
    pub text: String,
    pub hanging: Option<f32>,
    pub suff: String,
    pub tab_stop: Option<f32>,
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
        let resolved = self.resolve_marker_full(num_id, ilvl, counters);
        (resolved.text, resolved.hanging)
    }

    /// Resolve marker with full info (suff, tab_stop).
    pub fn resolve_marker_full(
        &self,
        num_id: &str,
        ilvl: u8,
        counters: &mut HashMap<(String, u8), u32>,
    ) -> ResolvedMarker {
        let fallback = ResolvedMarker {
            text: "\u{2022}".to_string(),
            hanging: Some(18.0),
            suff: "tab".to_string(),
            tab_stop: None,
        };

        let abstract_num_id = match self.num_map.get(num_id) {
            Some(id) => id,
            None => return fallback,
        };

        let abstract_num = match self.abstract_nums.get(abstract_num_id) {
            Some(an) => an,
            None => return fallback,
        };

        let level = match abstract_num.levels.get(&ilvl) {
            Some(l) => l,
            None => return fallback,
        };

        let hanging = level.indent_hanging;
        let suff = level.suff.clone();
        let tab_stop = level.tab_stop;

        if level.num_fmt == "bullet" {
            let marker = if level.lvl_text.is_empty() {
                "\u{2022}".to_string()
            } else {
                // Map Symbol font private use area characters to standard Unicode
                map_symbol_bullets(&level.lvl_text)
            };
            return ResolvedMarker { text: marker, hanging, suff, tab_stop };
        }

        // Numbered list: increment counter
        let key = (num_id.to_string(), ilvl);
        let count = counters.entry(key).or_insert_with(|| {
            if let Some(overrides) = self.level_overrides.get(num_id) {
                if let Some(&start) = overrides.get(&ilvl) {
                    return start - 1;
                }
            }
            0
        });
        *count += 1;
        let current = *count;

        let formatted_num = format_number(current, &level.num_fmt);

        let marker = if level.lvl_text.is_empty() {
            format!("{}.", formatted_num)
        } else {
            let placeholder = format!("%{}", ilvl + 1);
            level.lvl_text.replace(&placeholder, &formatted_num)
        };

        ResolvedMarker { text: marker, hanging, suff, tab_stop }
    }

    /// Get the left indent for a given numId and ilvl
    pub fn get_level_indent(&self, num_id: &str, ilvl: u8) -> Option<f32> {
        let abstract_num_id = self.num_map.get(num_id)?;
        let abstract_num = self.abstract_nums.get(abstract_num_id)?;
        let level = abstract_num.levels.get(&ilvl)?;
        level.indent_left
    }
}

/// Map Symbol/Wingdings private use area characters to standard Unicode equivalents
fn map_symbol_bullets(text: &str) -> String {
    text.chars().map(|ch| {
        match ch {
            '\u{F0B7}' => '\u{2022}', // Symbol bullet → • (bullet)
            '\u{F06F}' => '\u{25CB}', // Symbol circle → ○
            '\u{F0A7}' => '\u{25AA}', // Symbol square → ▪
            '\u{F0FC}' => '\u{2713}', // Wingdings checkmark → ✓
            '\u{F0D8}' => '\u{25B6}', // Symbol arrow → ▶
            '\u{F076}' => '\u{2756}', // Wingdings diamond → ◆ (approx)
            '\u{F0A8}' => '\u{25A0}', // Symbol filled square → ■
            '\u{F06E}' => '\u{25CF}', // Symbol filled circle → ●
            other => other,
        }
    }).collect()
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
                        let (abstract_num_id, overrides) = parse_num_element(&mut reader)?;
                        defs.num_map.insert(num_id.clone(), abstract_num_id);
                        if !overrides.is_empty() {
                            defs.level_overrides.insert(num_id, overrides);
                        }
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
    let mut suff = "tab".to_string();
    let mut tab_stop = None;
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
                    "suff" => {
                        for attr in e.attributes().flatten() {
                            if local_name(attr.key.as_ref()) == "val" {
                                suff = String::from_utf8_lossy(&attr.value).to_string();
                            }
                        }
                    }
                    "tab" => {
                        // w:tabs/w:tab — tab stop position for numbering
                        for attr in e.attributes().flatten() {
                            let key = local_name(attr.key.as_ref());
                            if key == "pos" {
                                let val = String::from_utf8_lossy(&attr.value);
                                tab_stop = val.parse::<f32>().ok().map(|v| v / 20.0);
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
        suff,
        tab_stop,
    })
}

/// Parse a w:num element to get its abstractNumId and any lvlOverride/startOverride
fn parse_num_element(reader: &mut Reader<&[u8]>) -> Result<(String, HashMap<u8, u32>), ParseError> {
    let mut abstract_num_id = String::new();
    let mut overrides: HashMap<u8, u32> = HashMap::new();
    let mut depth = 0;
    let mut current_override_ilvl: Option<u8> = None;

    loop {
        match reader.read_event()? {
            Event::Start(e) => {
                let local = local_name(e.name().as_ref());
                if local == "lvlOverride" {
                    for attr in e.attributes().flatten() {
                        if local_name(attr.key.as_ref()) == "ilvl" {
                            let val = String::from_utf8_lossy(&attr.value);
                            current_override_ilvl = val.parse().ok();
                        }
                    }
                }
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
                } else if local == "startOverride" {
                    if let Some(ilvl) = current_override_ilvl {
                        for attr in e.attributes().flatten() {
                            if local_name(attr.key.as_ref()) == "val" {
                                let val = String::from_utf8_lossy(&attr.value);
                                if let Ok(start) = val.parse::<u32>() {
                                    overrides.insert(ilvl, start);
                                }
                            }
                        }
                    }
                }
            }
            Event::End(e) => {
                let local = local_name(e.name().as_ref());
                if local == "lvlOverride" {
                    current_override_ilvl = None;
                }
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

    Ok((abstract_num_id, overrides))
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
