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
    /// Apply tint/shade using HSL color space (matches Word's behavior).
    /// Positive values = tint (lighten towards white).
    /// Negative values = shade (darken towards black).
    /// Word uses HSL-based shade/tint, not simple RGB multiplication.
    pub fn apply_tint_shade(hex: &str, tint_shade: f64) -> String {
        let r = u8::from_str_radix(&hex[0..2], 16).unwrap_or(0) as f64 / 255.0;
        let g = u8::from_str_radix(&hex[2..4], 16).unwrap_or(0) as f64 / 255.0;
        let b = u8::from_str_radix(&hex[4..6], 16).unwrap_or(0) as f64 / 255.0;

        // Convert RGB to HSL
        let max = r.max(g).max(b);
        let min = r.min(g).min(b);
        let l = (max + min) / 2.0;

        let (h, s) = if (max - min).abs() < 1e-10 {
            (0.0, 0.0)
        } else {
            let d = max - min;
            let s = if l > 0.5 { d / (2.0 - max - min) } else { d / (max + min) };
            let h = if (max - r).abs() < 1e-10 {
                (g - b) / d + if g < b { 6.0 } else { 0.0 }
            } else if (max - g).abs() < 1e-10 {
                (b - r) / d + 2.0
            } else {
                (r - g) / d + 4.0
            };
            (h / 6.0, s)
        };

        // Apply tint/shade to luminance
        let l2 = if tint_shade > 0.0 {
            l * (1.0 - tint_shade) + tint_shade  // tint: move L towards 1.0
        } else {
            l * (1.0 + tint_shade)  // shade: multiply L by factor
        };

        // Convert HSL back to RGB
        let hsl_to_rgb = |h: f64, s: f64, l: f64| -> (u8, u8, u8) {
            if s.abs() < 1e-10 {
                let v = (l * 255.0).round() as u8;
                return (v, v, v);
            }
            let q = if l < 0.5 { l * (1.0 + s) } else { l + s - l * s };
            let p = 2.0 * l - q;
            let hue_to_rgb = |p: f64, q: f64, mut t: f64| -> f64 {
                if t < 0.0 { t += 1.0; }
                if t > 1.0 { t -= 1.0; }
                if t < 1.0/6.0 { return p + (q - p) * 6.0 * t; }
                if t < 1.0/2.0 { return q; }
                if t < 2.0/3.0 { return p + (q - p) * (2.0/3.0 - t) * 6.0; }
                p
            };
            (
                (hue_to_rgb(p, q, h + 1.0/3.0) * 255.0).round() as u8,
                (hue_to_rgb(p, q, h) * 255.0).round() as u8,
                (hue_to_rgb(p, q, h - 1.0/3.0) * 255.0).round() as u8,
            )
        };

        let (r2, g2, b2) = hsl_to_rgb(h, s, l2);
        format!("{:02X}{:02X}{:02X}", r2, g2, b2)
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
    // S323 (2026-05-26): re-wire `ea_empty_*` flags. S245 removed these
    // because they were dead; d1e8ac8 evidence shows the suppression IS
    // needed. Word's COM Font.NameFarEast for d1e8ac8 para 6 (殿) returns
    // "ＭＳ 明朝" (from rPrDefault inheritance), NOT "游明朝" (theme's
    // Jpan-script font). So Word's behavior: <a:ea typeface=""/> means
    // "EA slot explicitly empty — do NOT use Jpan script as fallback,
    // fall through to rPrDefault instead". Env-gated to preserve baseline.
    let mut ea_empty_major = false;
    let mut ea_empty_minor = false;
    // S327 (2026-05-26): DEFAULT-ON. Explicit <a:ea typeface=""/> suppresses
    // the Jpan-script font fallback for that font slot. End-of-parse fallback
    // changes from "Meiryo" to "MS Mincho" for empty-ea slots (closer to
    // Word's COM-confirmed behavior for d1e8ac8). Full-corpus measurement:
    // Phase 2 mean_iou 0.9570 → 0.9592 (+0.0022 strict increase), 17 → 18
    // pass (≥0.99), Phase 1 53/55 unchanged. d1e8ac8 +0.0607, 1636d28e
    // +0.0497, others +0.005-0.014, only b837 -0.0068 (acceptable noise).
    // Env-var preserved as OPT-OUT (set to "0" to disable).
    let suppress_jpan_when_empty_ea = std::env::var("OXI_S323_EMPTY_EA_NO_JPAN")
        .map(|v| v != "0" && v != "false")
        .unwrap_or(true);

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
                                    // S323: record explicit-empty-ea
                                    ea_empty_major = true;
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
                                    // S323: record explicit-empty-ea
                                    ea_empty_minor = true;
                                }
                            }
                        }
                    }
                    // Script-specific fonts (e.g. <a:font script="Jpan" typeface="ＭＳ ゴシック"/>)
                    // Historically Oxi used Jpan as the EA fallback even when
                    // <a:ea typeface=""/> was explicit-empty (S245 comment said
                    // "matches Word output"). S323 (2026-05-26) d1e8ac8 evidence
                    // overturns that: Word COM Font.NameFarEast for para 6 (殿)
                    // returns "ＭＳ 明朝" (rPrDefault), not "游明朝" (theme Jpan).
                    // So under OXI_S323_EMPTY_EA_NO_JPAN=1, explicit-empty <a:ea>
                    // SUPPRESSES the Jpan fallback for that font slot.
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
                            // S492e (2026-06-03, Lever E) — for empty-ea MAJOR, USE the
                            // Jpan font: Word renders 1ec1's title/headings in the Jpan
                            // MS Gothic, NOT the rPrDefault Mincho (visually confirmed:
                            // Word title=sans-Gothic, Oxi=serif-Mincho; the major-EA
                            // theme token "+見出しのフォント" is in use). S323 derived its
                            // blanket suppression from d1e8ac8's BODY (minor font) where
                            // Word DID use Mincho — so the suppression is MINOR-scoped, not
                            // major (re-derive from richer input, don't stack exceptions).
                            // Env OXI_E_MAJOR_JPAN, default OFF (byte-identical) for the canary.
                            let s492e_major_jpan = std::env::var("OXI_E_MAJOR_JPAN").is_ok();
                            let suppress_major = suppress_jpan_when_empty_ea && ea_empty_major
                                && !s492e_major_jpan;
                            let suppress_minor = suppress_jpan_when_empty_ea && ea_empty_minor;
                            if in_major_font && theme.major_font_ea.is_none() && !suppress_major {
                                theme.major_font_ea = Some(typeface.clone());
                            }
                            if in_minor_font && theme.minor_font_ea.is_none() && !suppress_minor {
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
    //
    // S323c (2026-05-26 revised x2): under S323 ea_empty_*, change the
    // end-of-parse fallback from "Meiryo" to **"MS Mincho"** (which
    // matches Word's COM-confirmed behavior on d1e8ac8 para 6:
    // Font.NameFarEast="ＭＳ 明朝" when ea is explicit-empty).
    //
    // Default ea-empty (no env): keeps Meiryo (historical behavior).
    // With env-gate: uses "MS Mincho" — closer to Word's actual
    // rendering for explicit-empty-ea docs.
    let ea_empty_fallback = if suppress_jpan_when_empty_ea {
        "MS Mincho"
    } else {
        "Meiryo"
    };
    if theme.major_font_ea.is_none() {
        theme.major_font_ea = Some(
            if suppress_jpan_when_empty_ea && ea_empty_major {
                ea_empty_fallback
            } else {
                "Meiryo"
            }.to_string()
        );
    }
    if theme.minor_font_ea.is_none() {
        theme.minor_font_ea = Some(
            if suppress_jpan_when_empty_ea && ea_empty_minor {
                ea_empty_fallback
            } else {
                "Meiryo"
            }.to_string()
        );
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
