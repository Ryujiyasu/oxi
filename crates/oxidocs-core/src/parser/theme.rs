use std::collections::HashMap;
use quick_xml::events::Event;
use quick_xml::reader::Reader;

/// Theme color scheme parsed from theme1.xml
#[derive(Debug, Clone, Default)]
pub struct ThemeColors {
    /// Named theme colors: "dk1", "lt1", "dk2", "lt2", "accent1"-"accent6", "hlink", "folHlink"
    pub colors: HashMap<String, String>,
    /// Major font (headings)
    pub major_font: Option<String>,
    /// Minor font (body)
    pub minor_font: Option<String>,
    /// Major East Asian font
    pub major_font_ea: Option<String>,
    /// Minor East Asian font
    pub minor_font_ea: Option<String>,
}

impl ThemeColors {
    /// Resolve a themeColor name to an RGB hex string
    pub fn resolve(&self, theme_color: &str) -> Option<&String> {
        // Map Word's themeColor attribute names to theme XML names
        let key = match theme_color {
            "dark1" | "text1" => "dk1",
            "light1" | "background1" => "lt1",
            "dark2" | "text2" => "dk2",
            "light2" | "background2" => "lt2",
            "accent1" => "accent1",
            "accent2" => "accent2",
            "accent3" => "accent3",
            "accent4" => "accent4",
            "accent5" => "accent5",
            "accent6" => "accent6",
            "hyperlink" => "hlink",
            "followedHyperlink" => "folHlink",
            other => other,
        };
        self.colors.get(key)
    }

    /// Apply tint/shade transformation to a hex color
    pub fn apply_tint_shade(hex: &str, tint_shade: f64) -> String {
        let r = u8::from_str_radix(&hex[0..2], 16).unwrap_or(0);
        let g = u8::from_str_radix(&hex[2..4], 16).unwrap_or(0);
        let b = u8::from_str_radix(&hex[4..6], 16).unwrap_or(0);

        let (r, g, b) = if tint_shade > 0.0 {
            // Tint: move towards white
            let t = tint_shade;
            (
                (r as f64 + (255.0 - r as f64) * t) as u8,
                (g as f64 + (255.0 - g as f64) * t) as u8,
                (b as f64 + (255.0 - b as f64) * t) as u8,
            )
        } else {
            // Shade: move towards black
            let s = 1.0 + tint_shade;
            (
                (r as f64 * s) as u8,
                (g as f64 * s) as u8,
                (b as f64 * s) as u8,
            )
        };

        format!("{:02X}{:02X}{:02X}", r, g, b)
    }
}

/// Parse theme1.xml to extract color scheme and fonts
pub fn parse_theme(xml: &str) -> ThemeColors {
    let mut theme = ThemeColors::default();
    let mut reader = Reader::from_str(xml);

    let mut in_clr_scheme = false;
    let mut current_color_name: Option<String> = None;
    let mut in_major_font = false;
    let mut in_minor_font = false;
    // Track if <a:ea> was present but empty (typeface="") — suppresses script-based fallback
    let mut ea_empty_major = false;
    let mut ea_empty_minor = false;

    loop {
        match reader.read_event() {
            Ok(Event::Start(e)) => {
                let local = local_name(e.name().as_ref());
                match local.as_str() {
                    "clrScheme" => in_clr_scheme = true,
                    "majorFont" => in_major_font = true,
                    "minorFont" => in_minor_font = true,
                    // Theme color elements: dk1, lt1, dk2, lt2, accent1-6, hlink, folHlink
                    "dk1" | "lt1" | "dk2" | "lt2" | "accent1" | "accent2" | "accent3"
                    | "accent4" | "accent5" | "accent6" | "hlink" | "folHlink"
                        if in_clr_scheme =>
                    {
                        current_color_name = Some(local);
                    }
                    _ => {}
                }
            }
            Ok(Event::Empty(e)) => {
                let local = local_name(e.name().as_ref());
                match local.as_str() {
                    // Color value elements: srgbClr or sysClr
                    "srgbClr" if current_color_name.is_some() => {
                        for attr in e.attributes().flatten() {
                            if local_name(attr.key.as_ref()) == "val" {
                                let val = String::from_utf8_lossy(&attr.value).to_string();
                                if let Some(ref name) = current_color_name {
                                    theme.colors.insert(name.clone(), val);
                                }
                            }
                        }
                    }
                    "sysClr" if current_color_name.is_some() => {
                        // System color — use lastClr attribute as the actual value
                        for attr in e.attributes().flatten() {
                            if local_name(attr.key.as_ref()) == "lastClr" {
                                let val = String::from_utf8_lossy(&attr.value).to_string();
                                if let Some(ref name) = current_color_name {
                                    theme.colors.insert(name.clone(), val);
                                }
                            }
                        }
                    }
                    "latin" if in_major_font && theme.major_font.is_none() => {
                        for attr in e.attributes().flatten() {
                            if local_name(attr.key.as_ref()) == "typeface" {
                                theme.major_font =
                                    Some(String::from_utf8_lossy(&attr.value).to_string());
                            }
                        }
                    }
                    "latin" if in_minor_font && theme.minor_font.is_none() => {
                        for attr in e.attributes().flatten() {
                            if local_name(attr.key.as_ref()) == "typeface" {
                                theme.minor_font =
                                    Some(String::from_utf8_lossy(&attr.value).to_string());
                            }
                        }
                    }
                    "ea" if in_major_font && theme.major_font_ea.is_none() => {
                        for attr in e.attributes().flatten() {
                            if local_name(attr.key.as_ref()) == "typeface" {
                                let val = String::from_utf8_lossy(&attr.value).to_string();
                                if !val.is_empty() {
                                    theme.major_font_ea = Some(val);
                                } else {
                                    ea_empty_major = true; // ea="" suppresses Jpan fallback
                                }
                            }
                        }
                    }
                    "ea" if in_minor_font && theme.minor_font_ea.is_none() => {
                        for attr in e.attributes().flatten() {
                            if local_name(attr.key.as_ref()) == "typeface" {
                                let val = String::from_utf8_lossy(&attr.value).to_string();
                                if !val.is_empty() {
                                    theme.minor_font_ea = Some(val);
                                } else {
                                    ea_empty_minor = true;
                                }
                            }
                        }
                    }
                    // Script-specific fonts (e.g. <a:font script="Jpan" typeface="ＭＳ ゴシック"/>)
                    // Matches Word output: ea="" does NOT suppress Jpan font lookup.
                    // Word uses Jpan supplemental font even when ea typeface is empty.
                    "font" if (in_major_font || in_minor_font) => {
                        let mut script = String::new();
                        let mut typeface = String::new();
                        for attr in e.attributes().flatten() {
                            let key = local_name(attr.key.as_ref());
                            let val = String::from_utf8_lossy(&attr.value).to_string();
                            match key.as_str() {
                                "script" => script = val,
                                "typeface" => typeface = val,
                                _ => {}
                            }
                        }
                        if script == "Jpan" && !typeface.is_empty() {
                            if in_major_font && theme.major_font_ea.is_none() {
                                theme.major_font_ea = Some(typeface.clone());
                            }
                            if in_minor_font && theme.minor_font_ea.is_none() {
                                theme.minor_font_ea = Some(typeface);
                            }
                        }
                    }
                    _ => {}
                }
            }
            Ok(Event::End(e)) => {
                let local = local_name(e.name().as_ref());
                match local.as_str() {
                    "clrScheme" => in_clr_scheme = false,
                    "majorFont" => in_major_font = false,
                    "minorFont" => in_minor_font = false,
                    "dk1" | "lt1" | "dk2" | "lt2" | "accent1" | "accent2" | "accent3"
                    | "accent4" | "accent5" | "accent6" | "hlink" | "folHlink" => {
                        current_color_name = None;
                    }
                    _ => {}
                }
            }
            Ok(Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
    }

    // Fallback: when theme EA fonts are empty, use system defaults
    // Matches Word output: Japanese Windows uses Meiryo for headings when theme EA is unset
    if theme.major_font_ea.is_none() {
        theme.major_font_ea = Some("Meiryo".to_string());
    }
    if theme.minor_font_ea.is_none() {
        theme.minor_font_ea = Some("Meiryo".to_string());
    }

    theme
}

fn local_name(name: &[u8]) -> String {
    let s = std::str::from_utf8(name).unwrap_or("");
    match s.rfind(':') {
        Some(pos) => s[pos + 1..].to_string(),
        None => s.to_string(),
    }
}
