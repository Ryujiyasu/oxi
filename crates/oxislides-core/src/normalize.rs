// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! Resolve PresentationML's indirections in the XML itself, before the parser
//! reads it.
//!
//! Shapes that PowerPoint inserts with its own UI do not state their colours
//! directly. They point at the theme: `p:style` names a fill/line style from
//! the theme's `a:fmtScheme` and binds its placeholder colour `phClr`; colours
//! carry modifiers (`lumMod`, `lumOff`, `tint`, `shade`, ...) applied to a
//! scheme slot; the scheme slot itself (`bg1`, `tx1`, ...) goes through the
//! master's `p:clrMap`; backgrounds use `p:bgRef`; fonts use `+mn-lt`-style
//! theme tokens.
//!
//! The parser is an event loop with one arm per element and state, and it reads
//! literal values. Instead of threading every indirection through every arm,
//! each slide/layout/master part is rewritten so the indirections are already
//! resolved: style references become the literal `spPr` children they stand
//! for, every colour becomes an `a:srgbClr` (with an `a:alpha` child when it is
//! translucent), `bgRef` becomes `bgPr`, and font tokens become face names. The
//! parser then sees the same shape it sees for an explicitly-styled deck.
//!
//! `OXI_NORMALIZE_DISABLE` turns the whole pass off.

use std::collections::HashMap;

use oxidocs_common::archive::OoxmlArchive;
use oxidocs_common::relationships::parse_relationships;
use quick_xml::events::Event;
use quick_xml::reader::Reader;

// ---------------------------------------------------------------------------
// A minimal DOM that round-trips the raw (still-escaped) text and attributes.
// ---------------------------------------------------------------------------

#[derive(Clone, Debug)]
pub(crate) enum Node {
    Elem(Elem),
    /// Raw markup that is copied through verbatim: text (still escaped),
    /// comments, CDATA, declarations, processing instructions.
    Raw(String),
}

#[derive(Clone, Debug)]
pub(crate) struct Elem {
    /// Qualified name as written, e.g. `a:srgbClr`.
    pub name: String,
    /// (qualified name, raw escaped value)
    pub attrs: Vec<(String, String)>,
    pub children: Vec<Node>,
}

impl Elem {
    fn new(name: String) -> Self {
        Elem { name, attrs: Vec::new(), children: Vec::new() }
    }

    pub fn local(&self) -> &str {
        local(&self.name)
    }

    fn prefix(&self) -> Option<&str> {
        self.name.split_once(':').map(|(p, _)| p)
    }

    /// The same prefix as this element, applied to `local`.
    fn sibling_name(&self, local: &str) -> String {
        match self.prefix() {
            Some(p) => format!("{p}:{local}"),
            None => local.to_string(),
        }
    }

    pub fn attr(&self, key: &str) -> Option<&str> {
        self.attrs
            .iter()
            .find(|(k, _)| k == key || local(k) == key && !k.starts_with("xmlns"))
            .map(|(_, v)| v.as_str())
    }

    fn set_attr(&mut self, key: &str, val: String) {
        if let Some(a) = self.attrs.iter_mut().find(|(k, _)| k == key) {
            a.1 = val;
        } else {
            self.attrs.push((key.to_string(), val));
        }
    }

    fn child_elems(&self) -> impl Iterator<Item = &Elem> {
        self.children.iter().filter_map(|n| match n {
            Node::Elem(e) => Some(e),
            Node::Raw(_) => None,
        })
    }

    fn child(&self, local_name: &str) -> Option<&Elem> {
        self.child_elems().find(|e| e.local() == local_name)
    }

    fn child_mut(&mut self, local_name: &str) -> Option<&mut Elem> {
        self.children.iter_mut().find_map(|n| match n {
            Node::Elem(e) if e.local() == local_name => Some(e),
            _ => None,
        })
    }

    fn has_child_in(&self, names: &[&str]) -> bool {
        self.child_elems().any(|e| names.contains(&e.local()))
    }

    /// Index of the first child element whose local name is in `names`.
    fn position_of_first(&self, names: &[&str]) -> Option<usize> {
        self.children.iter().position(|n| match n {
            Node::Elem(e) => names.contains(&e.local()),
            Node::Raw(_) => false,
        })
    }

    /// Insert before the first child named in `before`, or at the end.
    fn insert_before(&mut self, node: Node, before: &[&str]) {
        let at = self.position_of_first(before).unwrap_or(self.children.len());
        self.children.insert(at, node);
    }
}

fn local(qname: &str) -> &str {
    qname.rsplit_once(':').map(|(_, l)| l).unwrap_or(qname)
}

pub(crate) fn parse_dom(xml: &str) -> Option<Vec<Node>> {
    let mut reader = Reader::from_str(xml);
    let cfg = reader.config_mut();
    cfg.trim_text(false);
    cfg.expand_empty_elements = false;
    let mut stack: Vec<Elem> = Vec::new();
    let mut top: Vec<Node> = Vec::new();

    fn elem_of(e: &quick_xml::events::BytesStart) -> Elem {
        let mut el = Elem::new(String::from_utf8_lossy(e.name().as_ref()).into_owned());
        for a in e.attributes().with_checks(false).flatten() {
            el.attrs.push((
                String::from_utf8_lossy(a.key.as_ref()).into_owned(),
                String::from_utf8_lossy(&a.value).into_owned(),
            ));
        }
        el
    }
    fn push(stack: &mut [Elem], top: &mut Vec<Node>, n: Node) {
        match stack.last_mut() {
            Some(p) => p.children.push(n),
            None => top.push(n),
        }
    }

    loop {
        match reader.read_event().ok()? {
            Event::Start(e) => stack.push(elem_of(&e)),
            Event::Empty(e) => {
                let el = elem_of(&e);
                push(&mut stack, &mut top, Node::Elem(el));
            }
            Event::End(_) => {
                let el = stack.pop()?;
                push(&mut stack, &mut top, Node::Elem(el));
            }
            Event::Text(t) => {
                let s = String::from_utf8_lossy(&t).into_owned();
                push(&mut stack, &mut top, Node::Raw(s));
            }
            Event::CData(t) => {
                let s = format!("<![CDATA[{}]]>", String::from_utf8_lossy(&t));
                push(&mut stack, &mut top, Node::Raw(s));
            }
            Event::Comment(t) => {
                let s = format!("<!--{}-->", String::from_utf8_lossy(&t));
                push(&mut stack, &mut top, Node::Raw(s));
            }
            Event::Decl(d) => {
                let s = format!("<?{}?>", String::from_utf8_lossy(&d));
                push(&mut stack, &mut top, Node::Raw(s));
            }
            Event::PI(p) => {
                let s = format!("<?{}?>", String::from_utf8_lossy(&p));
                push(&mut stack, &mut top, Node::Raw(s));
            }
            Event::DocType(d) => {
                let s = format!("<!DOCTYPE {}>", String::from_utf8_lossy(&d));
                push(&mut stack, &mut top, Node::Raw(s));
            }
            Event::Eof => break,
        }
    }
    if !stack.is_empty() {
        return None;
    }
    Some(top)
}

pub(crate) fn write_dom(nodes: &[Node]) -> String {
    let mut out = String::new();
    fn w(n: &Node, out: &mut String) {
        match n {
            Node::Raw(s) => out.push_str(s),
            Node::Elem(e) => {
                out.push('<');
                out.push_str(&e.name);
                for (k, v) in &e.attrs {
                    out.push(' ');
                    out.push_str(k);
                    out.push_str("=\"");
                    // Raw values are already escaped, except that a value read
                    // from a single-quoted attribute may hold a bare `"`.
                    out.push_str(&v.replace('"', "&quot;"));
                    out.push('"');
                }
                if e.children.is_empty() {
                    out.push_str("/>");
                } else {
                    out.push('>');
                    for c in &e.children {
                        w(c, out);
                    }
                    out.push_str("</");
                    out.push_str(&e.name);
                    out.push('>');
                }
            }
        }
    }
    for n in nodes {
        w(n, &mut out);
    }
    out
}

// ---------------------------------------------------------------------------
// Theme
// ---------------------------------------------------------------------------

#[derive(Clone, Debug, Default)]
pub(crate) struct ThemeInfo {
    /// clrScheme slot -> RRGGBB
    pub colors: HashMap<String, String>,
    pub fill_styles: Vec<Elem>,
    pub ln_styles: Vec<Elem>,
    pub bg_fill_styles: Vec<Elem>,
    /// (latin, ea, cs) of majorFont / minorFont
    pub major: (String, String, String),
    pub minor: (String, String, String),
}

fn find_desc<'a>(nodes: &'a [Node], local_name: &str) -> Option<&'a Elem> {
    for n in nodes {
        if let Node::Elem(e) = n {
            if e.local() == local_name {
                return Some(e);
            }
            if let Some(f) = find_desc(&e.children, local_name) {
                return Some(f);
            }
        }
    }
    None
}

pub(crate) fn parse_theme_info(xml: &str) -> Option<ThemeInfo> {
    let dom = parse_dom(xml)?;
    let mut t = ThemeInfo::default();
    if let Some(cs) = find_desc(&dom, "clrScheme") {
        for slot in cs.child_elems() {
            if let Some(c) = slot.child_elems().next() {
                let hex = match c.local() {
                    "srgbClr" => c.attr("val").map(str::to_string),
                    "sysClr" => c.attr("lastClr").map(str::to_string).or_else(|| {
                        c.attr("val").map(|v| sys_color(v).to_string())
                    }),
                    _ => None,
                };
                if let Some(h) = hex.filter(|h| h.len() == 6) {
                    t.colors.insert(slot.local().to_string(), h.to_ascii_uppercase());
                }
            }
        }
    }
    let list = |name: &str| -> Vec<Elem> {
        find_desc(&dom, name)
            .map(|l| l.child_elems().cloned().collect())
            .unwrap_or_default()
    };
    t.fill_styles = list("fillStyleLst");
    t.ln_styles = list("lnStyleLst");
    t.bg_fill_styles = list("bgFillStyleLst");
    let fonts = |name: &str| -> (String, String, String) {
        let f = find_desc(&dom, name);
        let face = |s: &str| {
            f.and_then(|f| f.child(s))
                .and_then(|e| e.attr("typeface"))
                .map(unescape_attr)
                .unwrap_or_default()
        };
        (face("latin"), face("ea"), face("cs"))
    };
    t.major = fonts("majorFont");
    t.minor = fonts("minorFont");
    Some(t)
}

fn unescape_attr(v: &str) -> String {
    quick_xml::escape::unescape(v).map(|c| c.into_owned()).unwrap_or_else(|_| v.to_string())
}

fn escape_attr(v: &str) -> String {
    quick_xml::escape::escape(v).into_owned()
}

/// `p:clrMap` / `a:overrideClrMapping` attributes: bg1 -> lt1 etc.
fn clr_map_of(e: &Elem) -> HashMap<String, String> {
    e.attrs
        .iter()
        .filter(|(k, _)| !k.starts_with("xmlns"))
        .map(|(k, v)| (local(k).to_string(), v.clone()))
        .collect()
}

fn default_clr_map() -> HashMap<String, String> {
    [("bg1", "lt1"), ("tx1", "dk1"), ("bg2", "lt2"), ("tx2", "dk2")]
        .iter()
        .map(|(a, b)| (a.to_string(), b.to_string()))
        .collect()
}

// ---------------------------------------------------------------------------
// Colour arithmetic
// ---------------------------------------------------------------------------

const COLOR_ELEMS: &[&str] = &["srgbClr", "schemeClr", "sysClr", "prstClr", "scrgbClr", "hslClr"];

fn sys_color(v: &str) -> &'static str {
    match v {
        "window" | "highlightText" | "btnHighlight" | "3dLight" | "menu" | "infoText" => {
            if v == "infoText" {
                "000000"
            } else {
                "FFFFFF"
            }
        }
        "windowText" | "menuText" | "btnText" | "captionText" | "3dDkShadow" => "000000",
        "btnFace" | "menuBar" => "F0F0F0",
        "btnShadow" => "A0A0A0",
        "grayText" => "6D6D6D",
        "highlight" | "hotLight" => "0078D7",
        _ => "000000",
    }
}

fn preset_color(v: &str) -> Option<&'static str> {
    // The ST_PresetColorVal names that occur in practice (CSS colour values).
    let v = v.to_ascii_lowercase();
    let v = v.strip_prefix("dk").map(|s| format!("dark{s}")).unwrap_or(v);
    let v = v.strip_prefix("lt").map(|s| format!("light{s}")).unwrap_or(v);
    let v = v.strip_prefix("med").map(|s| format!("medium{s}")).unwrap_or(v);
    Some(match v.as_str() {
        "black" => "000000",
        "white" => "FFFFFF",
        "red" => "FF0000",
        "green" => "008000",
        "lime" => "00FF00",
        "blue" => "0000FF",
        "yellow" => "FFFF00",
        "cyan" | "aqua" => "00FFFF",
        "magenta" | "fuchsia" => "FF00FF",
        "gray" | "grey" => "808080",
        "silver" => "C0C0C0",
        "maroon" => "800000",
        "navy" => "000080",
        "olive" => "808000",
        "purple" => "800080",
        "teal" => "008080",
        "orange" => "FFA500",
        "pink" => "FFC0CB",
        "brown" => "A52A2A",
        "gold" => "FFD700",
        "darkgray" | "darkgrey" => "A9A9A9",
        "lightgray" | "lightgrey" => "D3D3D3",
        "darkred" => "8B0000",
        "darkgreen" => "006400",
        "darkblue" => "00008B",
        "lightblue" => "ADD8E6",
        "lightgreen" => "90EE90",
        "lightyellow" => "FFFFE0",
        "darkorange" => "FF8C00",
        "darkviolet" => "9400D3",
        "violet" => "EE82EE",
        "indigo" => "4B0082",
        "beige" => "F5F5DC",
        "tan" => "D2B48C",
        "coral" => "FF7F50",
        "salmon" => "FA8072",
        "tomato" => "FF6347",
        "crimson" => "DC143C",
        "khaki" => "F0E68C",
        "skyblue" => "87CEEB",
        "steelblue" => "4682B4",
        "royalblue" => "4169E1",
        "dodgerblue" => "1E90FF",
        "slategray" | "slategrey" => "708090",
        "dimgray" | "dimgrey" => "696969",
        "whitesmoke" => "F5F5F5",
        "gainsboro" => "DCDCDC",
        "ivory" => "FFFFF0",
        "lavender" => "E6E6FA",
        "turquoise" => "40E0D0",
        "chocolate" => "D2691E",
        "firebrick" => "B22222",
        "forestgreen" => "228B22",
        "seagreen" => "2E8B57",
        "orchid" => "DA70D6",
        "plum" => "DDA0DD",
        "sienna" => "A0522D",
        "wheat" => "F5DEB3",
        "aliceblue" => "F0F8FF",
        "snow" => "FFFAFA",
        "honeydew" => "F0FFF0",
        "mintcream" => "F5FFFA",
        "azure" => "F0FFFF",
        "linen" => "FAF0E6",
        _ => return None,
    })
}

fn hex_to_rgb(h: &str) -> Option<[f64; 3]> {
    if h.len() != 6 {
        return None;
    }
    let p = |i: usize| u8::from_str_radix(&h[i..i + 2], 16).ok().map(|v| v as f64 / 255.0);
    Some([p(0)?, p(2)?, p(4)?])
}

fn rgb_to_hex(c: [f64; 3]) -> String {
    let q = |v: f64| (v.clamp(0.0, 1.0) * 255.0).round() as u8;
    format!("{:02X}{:02X}{:02X}", q(c[0]), q(c[1]), q(c[2]))
}

fn to_linear(v: f64) -> f64 {
    if v <= 0.04045 {
        v / 12.92
    } else {
        ((v + 0.055) / 1.055).powf(2.4)
    }
}

fn to_srgb(v: f64) -> f64 {
    let v = v.clamp(0.0, 1.0);
    if v <= 0.0031308 {
        v * 12.92
    } else {
        1.055 * v.powf(1.0 / 2.4) - 0.055
    }
}

fn rgb_to_hsl(c: [f64; 3]) -> [f64; 3] {
    let (r, g, b) = (c[0], c[1], c[2]);
    let max = r.max(g).max(b);
    let min = r.min(g).min(b);
    let l = (max + min) / 2.0;
    if (max - min).abs() < 1e-12 {
        return [0.0, 0.0, l];
    }
    let d = max - min;
    let s = if l > 0.5 { d / (2.0 - max - min) } else { d / (max + min) };
    let h = if max == r {
        (g - b) / d + if g < b { 6.0 } else { 0.0 }
    } else if max == g {
        (b - r) / d + 2.0
    } else {
        (r - g) / d + 4.0
    };
    [h * 60.0, s, l]
}

fn hsl_to_rgb(hsl: [f64; 3]) -> [f64; 3] {
    let (h, s, l) = (hsl[0].rem_euclid(360.0) / 360.0, hsl[1].clamp(0.0, 1.0), hsl[2].clamp(0.0, 1.0));
    if s == 0.0 {
        return [l, l, l];
    }
    let q = if l < 0.5 { l * (1.0 + s) } else { l + s - l * s };
    let p = 2.0 * l - q;
    let f = |mut t: f64| {
        if t < 0.0 {
            t += 1.0;
        }
        if t > 1.0 {
            t -= 1.0;
        }
        if t < 1.0 / 6.0 {
            p + (q - p) * 6.0 * t
        } else if t < 0.5 {
            q
        } else if t < 2.0 / 3.0 {
            p + (q - p) * (2.0 / 3.0 - t) * 6.0
        } else {
            p
        }
    };
    [f(h + 1.0 / 3.0), f(h), f(h - 1.0 / 3.0)]
}

/// What a colour element resolves to.
#[derive(Clone, Copy, Debug, PartialEq)]
pub(crate) struct Resolved {
    pub rgb: [f64; 3],
    /// 0..=1
    pub alpha: f64,
}

pub(crate) struct ColorCtx<'a> {
    pub theme: &'a ThemeInfo,
    pub clr_map: &'a HashMap<String, String>,
}

impl ColorCtx<'_> {
    fn scheme(&self, val: &str) -> Option<String> {
        let slot = self.clr_map.get(val).map(String::as_str).unwrap_or(val);
        if let Some(h) = self.theme.colors.get(slot) {
            return Some(h.clone());
        }
        // A theme without that slot: the Office default palette, as the parser
        // always used.
        Some(
            match slot {
                "lt1" => "FFFFFF",
                "lt2" => "E7E6E6",
                "dk1" => "000000",
                "dk2" => "44546A",
                "accent1" => "4472C4",
                "accent2" => "ED7D31",
                "accent3" => "A5A5A5",
                "accent4" => "FFC000",
                "accent5" => "5B9BD5",
                "accent6" => "70AD47",
                "hlink" => "0563C1",
                "folHlink" => "954F72",
                _ => return None,
            }
            .to_string(),
        )
    }

    /// Resolve a colour element (`srgbClr`, `schemeClr`, ...) with its
    /// modifier children applied in document order.
    pub fn resolve(&self, e: &Elem) -> Option<Resolved> {
        let pct = |v: Option<&str>| v.and_then(|v| v.trim_end_matches('%').parse::<f64>().ok());
        let mut rgb = match e.local() {
            "srgbClr" => hex_to_rgb(e.attr("val")?)?,
            "schemeClr" => hex_to_rgb(&self.scheme(e.attr("val")?)?)?,
            "sysClr" => match e.attr("lastClr") {
                Some(l) => hex_to_rgb(l)?,
                None => hex_to_rgb(sys_color(e.attr("val")?))?,
            },
            "prstClr" => hex_to_rgb(preset_color(e.attr("val")?)?)?,
            "scrgbClr" => {
                let c = |k: &str| pct(e.attr(k)).map(|v| to_srgb(v / 100000.0));
                [c("r")?, c("g")?, c("b")?]
            }
            "hslClr" => {
                let h = pct(e.attr("hue"))? / 60000.0;
                let s = pct(e.attr("sat"))? / 100000.0;
                let l = pct(e.attr("lum"))? / 100000.0;
                hsl_to_rgb([h, s, l])
            }
            _ => return None,
        };
        let mut alpha = 1.0f64;
        for m in e.child_elems() {
            let Some(v) = pct(m.attr("val")) else {
                // Valueless modifiers.
                match m.local() {
                    "inv" => rgb = [1.0 - rgb[0], 1.0 - rgb[1], 1.0 - rgb[2]],
                    "comp" => {
                        let mut h = rgb_to_hsl(rgb);
                        h[0] += 180.0;
                        rgb = hsl_to_rgb(h);
                    }
                    "gray" => {
                        let y = 0.299 * rgb[0] + 0.587 * rgb[1] + 0.114 * rgb[2];
                        rgb = [y, y, y];
                    }
                    "gamma" => rgb = rgb.map(to_srgb),
                    "invGamma" => rgb = rgb.map(to_linear),
                    _ => {}
                }
                continue;
            };
            let f = v / 100000.0;
            match m.local() {
                "alpha" => alpha = f.clamp(0.0, 1.0),
                "alphaMod" => alpha = (alpha * f).clamp(0.0, 1.0),
                "alphaOff" => alpha = (alpha + f).clamp(0.0, 1.0),
                // tint / shade work on linear light (an accent1 4472C4 under
                // shade 50% is PowerPoint's 2F528F outline).
                "shade" => {
                    rgb = rgb.map(|c| to_srgb(to_linear(c) * f.clamp(0.0, 1.0)));
                }
                "tint" => {
                    let f = f.clamp(0.0, 1.0);
                    rgb = rgb.map(|c| to_srgb(1.0 - (1.0 - to_linear(c)) * f));
                }
                "lumMod" | "lumOff" | "lum" | "satMod" | "satOff" | "sat" | "hueMod"
                | "hueOff" | "hue" => {
                    let mut h = rgb_to_hsl(rgb);
                    match m.local() {
                        "lumMod" => h[2] *= f,
                        "lumOff" => h[2] += f,
                        "lum" => h[2] = f,
                        "satMod" => h[1] *= f,
                        "satOff" => h[1] += f,
                        "sat" => h[1] = f,
                        // hue values are in 60000ths of a degree
                        "hueMod" => h[0] *= f,
                        "hueOff" => h[0] += v / 60000.0,
                        "hue" => h[0] = v / 60000.0,
                        _ => {}
                    }
                    h[1] = h[1].clamp(0.0, 1.0);
                    h[2] = h[2].clamp(0.0, 1.0);
                    rgb = hsl_to_rgb(h);
                }
                "red" | "green" | "blue" | "redMod" | "greenMod" | "blueMod" | "redOff"
                | "greenOff" | "blueOff" => {
                    let i = match &m.local()[..3] {
                        "red" => 0,
                        "gre" => 1,
                        _ => 2,
                    };
                    let lin = to_linear(rgb[i]);
                    let n = if m.local().ends_with("Mod") {
                        lin * f
                    } else if m.local().ends_with("Off") {
                        lin + f
                    } else {
                        f
                    };
                    rgb[i] = to_srgb(n);
                }
                _ => {}
            }
        }
        Some(Resolved { rgb, alpha })
    }
}

fn is_plain_srgb(e: &Elem) -> bool {
    e.local() == "srgbClr"
        && e.attr("val").is_some_and(|v| v.len() == 6)
        && e.child_elems().all(|c| c.local() == "alpha")
}

/// The literal `a:srgbClr` a resolved colour is written as.
fn srgb_elem(like: &Elem, r: Resolved) -> Elem {
    let mut out = Elem::new(like.sibling_name("srgbClr"));
    out.attrs.push(("val".to_string(), rgb_to_hex(r.rgb)));
    if r.alpha < 1.0 {
        let mut a = Elem::new(like.sibling_name("alpha"));
        a.attrs.push(("val".to_string(), format!("{}", (r.alpha * 100000.0).round() as i64)));
        out.children.push(Node::Elem(a));
    }
    out
}

/// Replace every colour element below `nodes` with its resolved literal.
fn bake_colors(nodes: &mut [Node], cx: &ColorCtx) {
    for n in nodes.iter_mut() {
        if let Node::Elem(e) = n {
            if COLOR_ELEMS.contains(&e.local()) {
                if is_plain_srgb(e) {
                    continue;
                }
                if let Some(r) = cx.resolve(e) {
                    *e = srgb_elem(e, r);
                }
                continue;
            }
            // A theme's clrScheme / a chart's own palette are left alone; only
            // slide/layout/master parts reach this pass anyway.
            if e.local() == "clrScheme" {
                continue;
            }
            bake_colors(&mut e.children, cx);
        }
    }
}

// ---------------------------------------------------------------------------
// Style references
// ---------------------------------------------------------------------------

/// Replace `schemeClr val="phClr"` in a theme style with the reference colour,
/// the reference's own modifiers first and then the style's.
fn bind_ph_clr(nodes: &mut [Node], ref_color: Option<&Elem>) {
    for n in nodes.iter_mut() {
        if let Node::Elem(e) = n {
            if e.local() == "schemeClr" && e.attr("val") == Some("phClr") {
                if let Some(rc) = ref_color {
                    let mut bound = rc.clone();
                    // the style's modifiers follow the reference's
                    bound.children.extend(e.children.iter().cloned());
                    *e = bound;
                }
                continue;
            }
            bind_ph_clr(&mut e.children, ref_color);
        }
    }
}

fn ref_color(r: &Elem) -> Option<&Elem> {
    r.child_elems().find(|c| COLOR_ELEMS.contains(&c.local()))
}

const FILL_ELEMS: &[&str] = &["noFill", "solidFill", "gradFill", "blipFill", "pattFill", "grpFill"];
const LN_FILL_ELEMS: &[&str] = &["noFill", "solidFill", "gradFill", "pattFill"];

fn style_fill(theme: &ThemeInfo, idx: u32, rc: Option<&Elem>) -> Option<Elem> {
    let base = match idx {
        0 => return None,
        1..=999 => theme.fill_styles.get(idx as usize - 1)?,
        _ => theme.bg_fill_styles.get(idx.checked_sub(1001)? as usize)?,
    };
    let mut f = base.clone();
    let mut v = vec![Node::Elem(f)];
    bind_ph_clr(&mut v, rc);
    f = match v.pop() {
        Some(Node::Elem(e)) => e,
        _ => return None,
    };
    Some(f)
}

/// Write a `p:style`'s fill/line into `spPr` where `spPr` does not state its
/// own, and its font colour into runs that do not state one.
fn apply_shape_style(shape: &mut Elem, cx: &ColorCtx) {
    let Some(style) = shape.child("style").cloned() else {
        return;
    };
    let idx = |name: &str| -> Option<(u32, Option<Elem>)> {
        let r = style.child(name)?;
        let i = r.attr("idx")?.parse::<u32>().ok()?;
        Some((i, ref_color(r).cloned()))
    };
    let fill_ref = idx("fillRef");
    let ln_ref = idx("lnRef");
    let font_ref = style.child("fontRef").cloned();

    if let Some(sp_pr) = shape.child_mut("spPr") {
        // Fill.
        if !sp_pr.has_child_in(FILL_ELEMS) {
            if let Some((i, rc)) = &fill_ref {
                if let Some(f) = style_fill(cx.theme, *i, rc.as_ref()) {
                    sp_pr.insert_before(
                        Node::Elem(f),
                        &["ln", "effectLst", "effectDag", "scene3d", "sp3d", "extLst"],
                    );
                }
            }
        }
        // Line.
        if let Some((i, rc)) = &ln_ref {
            if *i > 0 {
                if let Some(base) = cx.theme.ln_styles.get(*i as usize - 1) {
                    let mut theme_ln = base.clone();
                    let mut v = vec![Node::Elem(theme_ln)];
                    bind_ph_clr(&mut v, rc.as_ref());
                    theme_ln = match v.pop() {
                        Some(Node::Elem(e)) => e,
                        _ => unreachable!(),
                    };
                    match sp_pr.child_mut("ln") {
                        Some(own) => merge_ln(own, &theme_ln),
                        None => {
                            // Use the prefix of spPr's own children.
                            sp_pr.insert_before(
                                Node::Elem(theme_ln),
                                &["effectLst", "effectDag", "scene3d", "sp3d", "extLst"],
                            );
                        }
                    }
                }
            }
        }
    }

    // Font colour / face.
    if let Some(fr) = font_ref {
        let color = ref_color(&fr).cloned();
        let major = fr.attr("idx") == Some("major");
        if color.is_some() || major {
            if let Some(tx) = shape.child_mut("txBody") {
                if !lst_style_states_fill(tx) {
                    apply_font_ref(&mut tx.children, color.as_ref(), major);
                }
            }
        }
    }

    // The reference has been written out; drop it so no arm of the parser can
    // read its colours as something else.
    shape.children.retain(|n| !matches!(n, Node::Elem(e) if e.local() == "style"));
}

/// Fill in the line properties `own` leaves unstated from the theme line.
fn merge_ln(own: &mut Elem, theme_ln: &Elem) {
    // `<a:ln><a:noFill/></a:ln>` says "no outline": nothing of the theme's
    // line applies (its width would otherwise give the parser a stroke).
    if own.child("noFill").is_some() {
        return;
    }
    for (k, v) in &theme_ln.attrs {
        if !k.starts_with("xmlns") && own.attr(local(k)).is_none() {
            own.attrs.push((local(k).to_string(), v.clone()));
        }
    }
    let take = |names: &[&str]| -> Option<Node> {
        theme_ln
            .child_elems()
            .find(|c| names.contains(&c.local()))
            .cloned()
            .map(Node::Elem)
    };
    if !own.has_child_in(LN_FILL_ELEMS) {
        if let Some(f) = take(LN_FILL_ELEMS) {
            own.children.insert(0, f);
        }
    }
    if !own.has_child_in(&["prstDash", "custDash"]) {
        if let Some(d) = take(&["prstDash", "custDash"]) {
            own.insert_before(d, &["round", "bevel", "miter", "headEnd", "tailEnd", "extLst"]);
        }
    }
    if !own.has_child_in(&["round", "bevel", "miter"]) {
        if let Some(j) = take(&["round", "bevel", "miter"]) {
            own.insert_before(j, &["headEnd", "tailEnd", "extLst"]);
        }
    }
}

fn lst_style_states_fill(tx_body: &Elem) -> bool {
    fn any_fill(e: &Elem) -> bool {
        e.child_elems().any(|c| FILL_ELEMS.contains(&c.local()) || any_fill(c))
    }
    tx_body.child("lstStyle").is_some_and(any_fill)
}

const RPR_FILL_BEFORE: &[&str] = &[
    "effectLst", "effectDag", "highlight", "uLnTx", "uLn", "uFillTx", "uFill", "latin", "ea",
    "cs", "sym", "hlinkClick", "hlinkMouseOver", "rtl", "extLst",
];

fn apply_font_ref(nodes: &mut [Node], color: Option<&Elem>, major: bool) {
    for n in nodes.iter_mut() {
        let Node::Elem(e) = n else { continue };
        match e.local() {
            "r" | "fld" => {
                if e.child("rPr").is_none() {
                    let rpr = Elem::new(e.sibling_name("rPr"));
                    e.children.insert(0, Node::Elem(rpr));
                }
                if let Some(rpr) = e.child_mut("rPr") {
                    fill_rpr(rpr, color, major);
                }
            }
            "endParaRPr" => fill_rpr(e, color, major),
            _ => apply_font_ref(&mut e.children, color, major),
        }
    }
}

fn fill_rpr(rpr: &mut Elem, color: Option<&Elem>, major: bool) {
    if let Some(c) = color {
        if !rpr.has_child_in(FILL_ELEMS) {
            let mut sf = Elem::new(rpr.sibling_name("solidFill"));
            sf.children.push(Node::Elem(c.clone()));
            rpr.insert_before(Node::Elem(sf), RPR_FILL_BEFORE);
        }
    }
    if major && rpr.child("latin").is_none() {
        let mut l = Elem::new(rpr.sibling_name("latin"));
        l.attrs.push(("typeface".to_string(), "+mj-lt".to_string()));
        rpr.insert_before(Node::Elem(l), &["ea", "cs", "sym", "hlinkClick", "hlinkMouseOver", "rtl", "extLst"]);
    }
}

fn apply_styles(nodes: &mut [Node], cx: &ColorCtx) {
    for n in nodes.iter_mut() {
        if let Node::Elem(e) = n {
            if matches!(e.local(), "sp" | "cxnSp") {
                apply_shape_style(e, cx);
            }
            apply_styles(&mut e.children, cx);
        }
    }
}

/// `p:bg/p:bgRef idx=".."` -> `p:bg/p:bgPr` holding the theme fill.
fn apply_bg_ref(nodes: &mut [Node], cx: &ColorCtx) {
    for n in nodes.iter_mut() {
        if let Node::Elem(e) = n {
            if e.local() == "bg" {
                let pos = e.children.iter().position(
                    |c| matches!(c, Node::Elem(b) if b.local() == "bgRef"),
                );
                if let Some(pos) = pos {
                    let Node::Elem(bg_ref) = &e.children[pos] else { unreachable!() };
                    let idx = bg_ref.attr("idx").and_then(|v| v.parse::<u32>().ok()).unwrap_or(0);
                    let rc = ref_color(bg_ref).cloned();
                    let mut bg_pr = Elem::new(bg_ref.sibling_name("bgPr"));
                    let fill = style_fill(cx.theme, idx, rc.as_ref());
                    let a_prefixed = |l: &str| {
                        rc.as_ref().map(|c| c.sibling_name(l)).unwrap_or_else(|| format!("a:{l}"))
                    };
                    match fill {
                        Some(f) => bg_pr.children.push(Node::Elem(f)),
                        None => match rc.clone() {
                            // No style to stand on: the reference colour itself.
                            Some(c) if idx > 0 => {
                                let mut sf = Elem::new(a_prefixed("solidFill"));
                                sf.children.push(Node::Elem(c));
                                bg_pr.children.push(Node::Elem(sf));
                            }
                            _ => bg_pr.children.push(Node::Elem(Elem::new(a_prefixed("noFill")))),
                        },
                    }
                    bg_pr.children.push(Node::Elem(Elem::new(a_prefixed("effectLst"))));
                    e.children[pos] = Node::Elem(bg_pr);
                }
                continue;
            }
            apply_bg_ref(&mut e.children, cx);
        }
    }
}

/// `+mn-lt`, `+mj-ea`, ... -> the theme's face. A token whose face is empty
/// in the theme drops the element, so the next font in the chain applies.
fn apply_font_tokens(nodes: &mut Vec<Node>, theme: &ThemeInfo) {
    nodes.retain_mut(|n| {
        let Node::Elem(e) = n else { return true };
        if matches!(e.local(), "latin" | "ea" | "cs" | "font" | "buFont") {
            if let Some(tf) = e.attr("typeface").map(str::to_string) {
                let face = match tf.as_str() {
                    "+mj-lt" => Some(&theme.major.0),
                    "+mj-ea" => Some(&theme.major.1),
                    "+mj-cs" => Some(&theme.major.2),
                    "+mn-lt" => Some(&theme.minor.0),
                    "+mn-ea" => Some(&theme.minor.1),
                    "+mn-cs" => Some(&theme.minor.2),
                    _ => None,
                };
                if let Some(face) = face {
                    if face.is_empty() {
                        return false;
                    }
                    e.set_attr("typeface", escape_attr(face));
                }
            }
        }
        apply_font_tokens(&mut e.children, theme);
        true
    });
}

// ---------------------------------------------------------------------------
// Table styles
// ---------------------------------------------------------------------------

/// PowerPoint's default for a new table, "Medium Style 2 - Accent 1", and its
/// siblings. PowerPoint writes the definition of every style a deck uses into
/// ppt/tableStyles.xml, but other producers often name the GUID alone.
fn builtin_table_style(id: &str) -> Option<String> {
    let accent = match id.to_ascii_uppercase().as_str() {
        "{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}" => "accent1",
        "{21E4AEA4-8DFA-4A89-87EB-49C32662AFE8}" => "accent2",
        "{F5AB1C69-6EDB-4FF4-983F-18BD219EF322}" => "accent3",
        "{00A15C55-8517-42AA-B614-E9B94910E393}" => "accent4",
        "{7DF18680-E054-41AD-8BC1-D1AEF772440D}" => "accent5",
        "{93296810-A885-4BE3-A3E7-6D5BEEA58F35}" => "accent6",
        "{073A0DAA-6AF3-43AB-8588-CEC1D06C72B9}" => "dk1",
        _ => return None,
    };
    let ln = |w: u32| {
        format!(r#"<a:ln w="{w}" cmpd="sng"><a:solidFill><a:schemeClr val="lt1"/></a:solidFill></a:ln>"#)
    };
    let l1 = ln(12700);
    let fill = |tint: Option<u32>| match tint {
        Some(t) => format!(
            r#"<a:fill><a:solidFill><a:schemeClr val="{accent}"><a:tint val="{t}"/></a:schemeClr></a:solidFill></a:fill>"#
        ),
        None => format!(r#"<a:fill><a:solidFill><a:schemeClr val="{accent}"/></a:solidFill></a:fill>"#),
    };
    let bold_lt1 = r#"<a:tcTxStyle b="on"><a:fontRef idx="minor"><a:prstClr val="black"/></a:fontRef><a:schemeClr val="lt1"/></a:tcTxStyle>"#;
    Some(format!(
        concat!(
            r#"<a:tblStyle xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" styleId="{id}">"#,
            r#"<a:wholeTbl><a:tcTxStyle><a:fontRef idx="minor"><a:prstClr val="black"/></a:fontRef><a:schemeClr val="dk1"/></a:tcTxStyle>"#,
            r#"<a:tcStyle><a:tcBdr><a:left>{l1}</a:left><a:right>{l1}</a:right><a:top>{l1}</a:top><a:bottom>{l1}</a:bottom><a:insideH>{l1}</a:insideH><a:insideV>{l1}</a:insideV></a:tcBdr>{f20}</a:tcStyle></a:wholeTbl>"#,
            r#"<a:band1H><a:tcStyle><a:tcBdr/>{f40}</a:tcStyle></a:band1H><a:band2H><a:tcStyle><a:tcBdr/></a:tcStyle></a:band2H>"#,
            r#"<a:band1V><a:tcStyle><a:tcBdr/>{f40}</a:tcStyle></a:band1V><a:band2V><a:tcStyle><a:tcBdr/></a:tcStyle></a:band2V>"#,
            r#"<a:lastCol>{b}<a:tcStyle><a:tcBdr/>{f}</a:tcStyle></a:lastCol>"#,
            r#"<a:firstCol>{b}<a:tcStyle><a:tcBdr/>{f}</a:tcStyle></a:firstCol>"#,
            r#"<a:lastRow>{b}<a:tcStyle><a:tcBdr><a:top>{l3}</a:top></a:tcBdr>{f}</a:tcStyle></a:lastRow>"#,
            r#"<a:firstRow>{b}<a:tcStyle><a:tcBdr><a:bottom>{l3}</a:bottom></a:tcBdr>{f}</a:tcStyle></a:firstRow>"#,
            r#"</a:tblStyle>"#
        ),
        id = id,
        l1 = l1,
        l3 = ln(38100),
        f20 = fill(Some(20000)),
        f40 = fill(Some(40000)),
        f = fill(None),
        b = bold_lt1,
    ))
}

/// styleId -> `a:tblStyle`, from ppt/tableStyles.xml.
pub(crate) fn parse_table_styles(xml: &str) -> HashMap<String, Elem> {
    let mut out = HashMap::new();
    if let Some(dom) = parse_dom(xml) {
        fn walk(nodes: &[Node], out: &mut HashMap<String, Elem>) {
            for n in nodes {
                if let Node::Elem(e) = n {
                    if e.local() == "tblStyle" {
                        if let Some(id) = e.attr("styleId") {
                            out.insert(id.to_ascii_uppercase(), e.clone());
                        }
                    } else {
                        walk(&e.children, out);
                    }
                }
            }
        }
        walk(&dom, &mut out);
    }
    out
}

#[derive(Default, Clone)]
struct CellLook {
    /// left, right, top, bottom: the `a:ln` (or resolved theme line)
    borders: [Option<Elem>; 4],
    fill: Option<Elem>,
    bold: Option<bool>,
    italic: Option<bool>,
    color: Option<Elem>,
}

/// Overlay one table-style part (`wholeTbl`, `firstRow`, ...) onto a cell.
/// `region` is the (r0, r1, c0, c1) block the part covers, which decides
/// whether a cell edge is the part's outer edge or one of its inside lines.
fn overlay_part(
    look: &mut CellLook,
    part: &Elem,
    cell: (usize, usize, usize),
    region: (usize, usize, usize, usize),
    cx: &ColorCtx,
) {
    let (r, c0c, c1c) = cell;
    let (rr0, rr1, rc0, rc1) = region;
    if let Some(tx) = part.child("tcTxStyle") {
        match tx.attr("b") {
            Some("on") => look.bold = Some(true),
            Some("off") => look.bold = Some(false),
            _ => {}
        }
        match tx.attr("i") {
            Some("on") => look.italic = Some(true),
            Some("off") => look.italic = Some(false),
            _ => {}
        }
        // A direct colour wins over the fontRef's.
        let direct = tx.child_elems().find(|c| COLOR_ELEMS.contains(&c.local()));
        let via_ref = tx.child("fontRef").and_then(ref_color);
        if let Some(col) = direct.or(via_ref) {
            look.color = Some(col.clone());
        }
    }
    let Some(st) = part.child("tcStyle") else { return };
    if let Some(bdr) = st.child("tcBdr") {
        let pick = |outer: bool, out_name: &str, in_name: &str| -> Option<String> {
            Some(if outer { out_name } else { in_name }.to_string())
        };
        let sides = [
            pick(c0c == rc0, "left", "insideV"),
            pick(c1c == rc1, "right", "insideV"),
            pick(r == rr0, "top", "insideH"),
            pick(r == rr1, "bottom", "insideH"),
        ];
        for (i, s) in sides.iter().enumerate() {
            let Some(s) = s else { continue };
            let Some(side) = bdr.child(s) else { continue };
            if let Some(ln) = side.child("ln") {
                look.borders[i] = Some(ln.clone());
            } else if let Some(lr) = side.child("lnRef") {
                let idx = lr.attr("idx").and_then(|v| v.parse::<usize>().ok()).unwrap_or(0);
                if idx == 0 {
                    look.borders[i] = None;
                } else if let Some(base) = cx.theme.ln_styles.get(idx - 1) {
                    let mut v = vec![Node::Elem(base.clone())];
                    bind_ph_clr(&mut v, ref_color(lr));
                    if let Some(Node::Elem(e)) = v.pop() {
                        look.borders[i] = Some(e);
                    }
                }
            }
        }
    }
    if let Some(f) = st.child("fill") {
        if let Some(inner) = f.child_elems().next() {
            look.fill = Some(inner.clone());
        }
    } else if let Some(fr) = st.child("fillRef") {
        let idx = fr.attr("idx").and_then(|v| v.parse::<u32>().ok()).unwrap_or(0);
        look.fill = style_fill(cx.theme, idx, ref_color(fr));
    }
}

fn apply_table_style(tbl: &mut Elem, styles: &HashMap<String, Elem>, cx: &ColorCtx) {
    let Some(tbl_pr) = tbl.child("tblPr") else { return };
    let flag = |k: &str| matches!(tbl_pr.attr(k), Some("1") | Some("true"));
    let (first_row, last_row, first_col, last_col, band_row, band_col) = (
        flag("firstRow"),
        flag("lastRow"),
        flag("firstCol"),
        flag("lastCol"),
        flag("bandRow"),
        flag("bandCol"),
    );
    let style: Option<Elem> = if let Some(inline) = tbl_pr.child("tableStyle") {
        Some(inline.clone())
    } else {
        let id = tbl_pr
            .child("tableStyleId")
            .and_then(|e| {
                e.children.iter().find_map(|n| match n {
                    Node::Raw(s) => Some(s.trim().to_ascii_uppercase()),
                    _ => None,
                })
            });
        id.and_then(|id| {
            styles.get(&id).cloned().or_else(|| {
                builtin_table_style(&id)
                    .and_then(|x| parse_dom(&x))
                    .and_then(|d| d.into_iter().find_map(|n| match n {
                        Node::Elem(e) => Some(e),
                        _ => None,
                    }))
            })
        })
    };
    let Some(style) = style else { return };

    let n_cols = tbl.child("tblGrid").map(|g| g.child_elems().count()).unwrap_or(0);
    let n_rows = tbl.child_elems().filter(|e| e.local() == "tr").count();
    if n_cols == 0 || n_rows == 0 {
        return;
    }
    let (lr, lc) = (n_rows - 1, n_cols - 1);

    let mut r = 0usize;
    for tr in tbl.children.iter_mut() {
        let Node::Elem(tr) = tr else { continue };
        if tr.local() != "tr" {
            continue;
        }
        let mut c = 0usize;
        for tc in tr.children.iter_mut() {
            let Node::Elem(tc) = tc else { continue };
            if tc.local() != "tc" {
                continue;
            }
            let span = tc.attr("gridSpan").and_then(|v| v.parse::<usize>().ok()).unwrap_or(1).max(1);
            let (c0, c1) = (c, (c + span - 1).min(lc));
            let mut look = CellLook::default();
            let whole = (0, lr, 0, lc);
            let row_region = (r, r, 0, lc);
            let col_region = |cc: usize| (0, lr, cc, cc);
            let cell = (r, c0, c1);
            if let Some(p) = style.child("wholeTbl") {
                overlay_part(&mut look, p, cell, whole, cx);
            }
            if band_row {
                let skip = (first_row && r == 0) || (last_row && r == lr);
                if !skip {
                    let k = r - first_row as usize;
                    let name = if k.is_multiple_of(2) { "band1H" } else { "band2H" };
                    if let Some(p) = style.child(name) {
                        overlay_part(&mut look, p, cell, row_region, cx);
                    }
                }
            }
            if band_col {
                let skip = (first_col && c0 == 0) || (last_col && c1 == lc);
                if !skip {
                    let k = c0 - first_col as usize;
                    let name = if k.is_multiple_of(2) { "band1V" } else { "band2V" };
                    if let Some(p) = style.child(name) {
                        overlay_part(&mut look, p, cell, col_region(c0), cx);
                    }
                }
            }
            if last_col && c1 == lc {
                if let Some(p) = style.child("lastCol") {
                    overlay_part(&mut look, p, cell, col_region(lc), cx);
                }
            }
            if first_col && c0 == 0 {
                if let Some(p) = style.child("firstCol") {
                    overlay_part(&mut look, p, cell, col_region(0), cx);
                }
            }
            if last_row && r == lr {
                if let Some(p) = style.child("lastRow") {
                    overlay_part(&mut look, p, cell, row_region, cx);
                }
            }
            if first_row && r == 0 {
                if let Some(p) = style.child("firstRow") {
                    overlay_part(&mut look, p, cell, row_region, cx);
                }
            }
            let corner = |name: &str| style.child(name);
            let single = (r, r, c0, c1);
            for (on, name) in [
                (first_row && last_col && r == 0 && c1 == lc, "neCell"),
                (first_row && first_col && r == 0 && c0 == 0, "nwCell"),
                (last_row && last_col && r == lr && c1 == lc, "seCell"),
                (last_row && first_col && r == lr && c0 == 0, "swCell"),
            ] {
                if on {
                    if let Some(p) = corner(name) {
                        overlay_part(&mut look, p, cell, single, cx);
                    }
                }
            }
            write_cell_look(tc, &look);
            c += span;
        }
        r += 1;
    }
}

fn write_cell_look(tc: &mut Elem, look: &CellLook) {
    if tc.child("tcPr").is_none() {
        let pr = Elem::new(tc.sibling_name("tcPr"));
        tc.insert_before(Node::Elem(pr), &["extLst"]);
    }
    let tc_pr = tc.child_mut("tcPr").expect("inserted above");
    let names = ["lnL", "lnR", "lnT", "lnB"];
    for (i, name) in names.iter().enumerate() {
        if tc_pr.child(name).is_some() {
            continue;
        }
        if let Some(ln) = &look.borders[i] {
            let mut e = ln.clone();
            e.name = tc_pr.sibling_name(name);
            let after: &[&str] = match i {
                0 => &["lnR", "lnT", "lnB", "lnTlToBr", "lnBlToTr", "cell3D", "noFill", "solidFill", "gradFill", "blipFill", "pattFill", "grpFill", "headers", "extLst"],
                1 => &["lnT", "lnB", "lnTlToBr", "lnBlToTr", "cell3D", "noFill", "solidFill", "gradFill", "blipFill", "pattFill", "grpFill", "headers", "extLst"],
                2 => &["lnB", "lnTlToBr", "lnBlToTr", "cell3D", "noFill", "solidFill", "gradFill", "blipFill", "pattFill", "grpFill", "headers", "extLst"],
                _ => &["lnTlToBr", "lnBlToTr", "cell3D", "noFill", "solidFill", "gradFill", "blipFill", "pattFill", "grpFill", "headers", "extLst"],
            };
            tc_pr.insert_before(Node::Elem(e), after);
        }
    }
    if !tc_pr.has_child_in(FILL_ELEMS) {
        if let Some(f) = &look.fill {
            tc_pr.insert_before(Node::Elem(f.clone()), &["headers", "extLst"]);
        }
    }
    if look.color.is_none() && look.bold.is_none() && look.italic.is_none() {
        return;
    }
    if let Some(tx) = tc.child_mut("txBody") {
        style_cell_runs(&mut tx.children, look);
    }
}

fn style_cell_runs(nodes: &mut [Node], look: &CellLook) {
    for n in nodes.iter_mut() {
        let Node::Elem(e) = n else { continue };
        match e.local() {
            "r" | "fld" => {
                if e.child("rPr").is_none() {
                    let rpr = Elem::new(e.sibling_name("rPr"));
                    e.children.insert(0, Node::Elem(rpr));
                }
                let rpr = e.child_mut("rPr").expect("inserted above");
                cell_rpr(rpr, look);
            }
            "endParaRPr" => cell_rpr(e, look),
            _ => style_cell_runs(&mut e.children, look),
        }
    }
}

fn cell_rpr(rpr: &mut Elem, look: &CellLook) {
    if let Some(b) = look.bold {
        if rpr.attr("b").is_none() {
            rpr.attrs.push(("b".into(), if b { "1" } else { "0" }.into()));
        }
    }
    if let Some(i) = look.italic {
        if rpr.attr("i").is_none() {
            rpr.attrs.push(("i".into(), if i { "1" } else { "0" }.into()));
        }
    }
    if let Some(c) = &look.color {
        if !rpr.has_child_in(FILL_ELEMS) {
            let mut sf = Elem::new(rpr.sibling_name("solidFill"));
            sf.children.push(Node::Elem(c.clone()));
            rpr.insert_before(Node::Elem(sf), RPR_FILL_BEFORE);
        }
    }
}

fn apply_tables(nodes: &mut [Node], styles: &HashMap<String, Elem>, cx: &ColorCtx) {
    for n in nodes.iter_mut() {
        if let Node::Elem(e) = n {
            if e.local() == "tbl" {
                apply_table_style(e, styles, cx);
                continue;
            }
            apply_tables(&mut e.children, styles, cx);
        }
    }
}

/// Rewrite one slide / layout / master part.
pub(crate) fn normalize_part(
    xml: &str,
    theme: &ThemeInfo,
    clr_map: &HashMap<String, String>,
) -> Option<String> {
    normalize_part_with_tables(xml, theme, clr_map, &HashMap::new())
}

pub(crate) fn normalize_part_with_tables(
    xml: &str,
    theme: &ThemeInfo,
    clr_map: &HashMap<String, String>,
    table_styles: &HashMap<String, Elem>,
) -> Option<String> {
    let mut dom = parse_dom(xml)?;
    let cx = ColorCtx { theme, clr_map };
    apply_styles(&mut dom, &cx);
    apply_tables(&mut dom, table_styles, &cx);
    apply_bg_ref(&mut dom, &cx);
    bake_colors(&mut dom, &cx);
    apply_font_tokens(&mut dom, theme);
    Some(write_dom(&dom))
}

/// The part's own colour-map override, if it states one.
fn own_clr_map(dom: &[Node]) -> Option<HashMap<String, String>> {
    let ovr = find_desc(dom, "clrMapOvr")?;
    let o = ovr.child("overrideClrMapping")?;
    Some(clr_map_of(o))
}

fn rels_path_of(part: &str) -> String {
    let cut = part.rfind('/').map(|i| i + 1).unwrap_or(0);
    format!("{}_rels/{}.rels", &part[..cut], &part[cut..])
}

fn resolve_target(from_part: &str, target: &str) -> String {
    if let Some(abs) = target.strip_prefix('/') {
        return abs.to_string();
    }
    let mut segs: Vec<&str> = from_part.split('/').collect();
    segs.pop();
    for s in target.split('/') {
        match s {
            ".." => {
                segs.pop();
            }
            "." | "" => {}
            s => segs.push(s),
        }
    }
    segs.join("/")
}

fn related(archive: &mut OoxmlArchive, part: &str, kind: &str) -> Option<String> {
    let rels_xml = archive.try_read_part(&rels_path_of(part)).ok()??;
    let rels = parse_relationships(&rels_xml).ok()?;
    let r = rels.values().find(|r| r.rel_type.ends_with(kind))?;
    Some(resolve_target(part, &r.target))
}

/// Normalise every slide, layout and master of the deck in place.
pub(crate) fn normalize_archive(archive: &mut OoxmlArchive) {
    if std::env::var("OXI_NORMALIZE_DISABLE").is_ok() {
        return;
    }
    let names = archive.file_names();
    let table_styles = archive
        .try_read_part("ppt/tableStyles.xml")
        .ok()
        .flatten()
        .map(|x| parse_table_styles(&x))
        .unwrap_or_default();
    let is_part = |dir: &str, n: &str| {
        n.starts_with(dir) && n.ends_with(".xml") && !n[dir.len()..].contains('/')
    };

    // master path -> (theme, clrMap)
    let mut masters: HashMap<String, (ThemeInfo, HashMap<String, String>)> = HashMap::new();
    for m in names.iter().filter(|n| is_part("ppt/slideMasters/", n)) {
        let Ok(Some(xml)) = archive.try_read_part(m) else { continue };
        let theme = related(archive, m, "/theme")
            .and_then(|t| archive.try_read_part(&t).ok().flatten())
            .and_then(|tx| parse_theme_info(&tx))
            .unwrap_or_default();
        let clr_map = parse_dom(&xml)
            .and_then(|d| find_desc(&d, "clrMap").map(clr_map_of))
            .unwrap_or_else(default_clr_map);
        if let Some(out) = normalize_part(&xml, &theme, &clr_map) {
            archive.set_part_override(m, out);
        }
        masters.insert(m.clone(), (theme, clr_map));
    }

    let mut layout_master: HashMap<String, String> = HashMap::new();
    for l in names.iter().filter(|n| is_part("ppt/slideLayouts/", n)) {
        let Some(mp) = related(archive, l, "/slideMaster") else { continue };
        layout_master.insert(l.clone(), mp.clone());
        let Some((theme, master_map)) = masters.get(&mp) else { continue };
        let Ok(Some(xml)) = archive.try_read_part(l) else { continue };
        let map = parse_dom(&xml).and_then(|d| own_clr_map(&d)).unwrap_or_else(|| master_map.clone());
        if let Some(out) = normalize_part(&xml, theme, &map) {
            archive.set_part_override(l, out);
        }
    }

    // Chart parts: their series colours are theme references too. A chart
    // has no master of its own; the deck's first master's theme stands in.
    let first_master = names
        .iter()
        .filter(|n| is_part("ppt/slideMasters/", n))
        .min_by_key(|n| {
            n.trim_start_matches("ppt/slideMasters/slideMaster")
                .trim_end_matches(".xml")
                .parse::<u32>()
                .unwrap_or(u32::MAX)
        })
        .and_then(|m| masters.get(m));
    if let Some((theme, map)) = first_master {
        for c in names.iter().filter(|n| is_part("ppt/charts/", n)) {
            let Ok(Some(xml)) = archive.try_read_part(c) else { continue };
            if let Some(out) = normalize_part(&xml, theme, map) {
                archive.set_part_override(c, out);
            }
        }
    }

    for s in names.iter().filter(|n| is_part("ppt/slides/", n)) {
        let Some(lp) = related(archive, s, "/slideLayout") else { continue };
        let Some(mp) = layout_master.get(&lp) else { continue };
        let Some((theme, master_map)) = masters.get(mp) else { continue };
        let Ok(Some(xml)) = archive.try_read_part(s) else { continue };
        let map = parse_dom(&xml).and_then(|d| own_clr_map(&d)).unwrap_or_else(|| master_map.clone());
        if let Some(out) = normalize_part_with_tables(&xml, theme, &map, &table_styles) {
            archive.set_part_override(s, out);
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn theme() -> ThemeInfo {
        let xml = r#"<a:theme xmlns:a="x"><a:themeElements><a:clrScheme name="Office">
<a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>
<a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>
<a:dk2><a:srgbClr val="44546A"/></a:dk2><a:lt2><a:srgbClr val="E7E6E6"/></a:lt2>
<a:accent1><a:srgbClr val="4472C4"/></a:accent1><a:accent2><a:srgbClr val="ED7D31"/></a:accent2>
</a:clrScheme>
<a:fontScheme name="Office"><a:majorFont><a:latin typeface="Calibri Light"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont>
<a:minorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont></a:fontScheme>
<a:fmtScheme name="Office"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
<a:gradFill><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="50000"/></a:schemeClr></a:gs></a:gsLst></a:gradFill>
<a:gradFill/></a:fillStyleLst>
<a:lnStyleLst><a:ln w="6350" cap="flat"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln>
<a:ln w="12700"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln><a:ln w="19050"/></a:lnStyleLst>
<a:effectStyleLst/><a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"><a:tint val="95000"/></a:schemeClr></a:solidFill><a:gradFill/></a:bgFillStyleLst>
</a:fmtScheme></a:themeElements></a:theme>"#;
        parse_theme_info(xml).unwrap()
    }

    fn norm(xml: &str) -> String {
        normalize_part(xml, &theme(), &default_clr_map()).unwrap()
    }

    #[test]
    fn default_inserted_shape_gets_theme_fill_line_and_white_text() {
        let xml = r#"<p:sp><p:spPr><a:xfrm/><a:prstGeom prst="rect"/></p:spPr><p:style><a:lnRef idx="2"><a:schemeClr val="accent1"><a:shade val="50000"/></a:schemeClr></a:lnRef><a:fillRef idx="1"><a:schemeClr val="accent1"/></a:fillRef><a:effectRef idx="0"><a:schemeClr val="accent1"/></a:effectRef><a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef></p:style><p:txBody><a:bodyPr/><a:p><a:r><a:t>Hi</a:t></a:r></a:p></p:txBody></p:sp>"#;
        let out = norm(xml);
        assert!(out.contains(r#"<a:prstGeom prst="rect"/><a:solidFill><a:srgbClr val="4472C4"/></a:solidFill><a:ln w="12700"><a:solidFill><a:srgbClr val="2F528F"/></a:solidFill>"#), "{out}");
        assert!(out.contains(r#"<a:r><a:rPr><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a:rPr><a:t>Hi</a:t></a:r>"#), "{out}");
        assert!(!out.contains("p:style"), "{out}");
    }

    #[test]
    fn own_fill_and_line_win_over_style() {
        let xml = r#"<p:sp><p:spPr><a:noFill/><a:ln w="38100"><a:prstDash val="dash"/></a:ln></p:spPr><p:style><a:lnRef idx="2"><a:schemeClr val="accent2"/></a:lnRef><a:fillRef idx="1"><a:schemeClr val="accent1"/></a:fillRef></p:style></p:sp>"#;
        let out = norm(xml);
        assert!(out.contains(r#"<a:noFill/><a:ln w="38100"><a:solidFill><a:srgbClr val="ED7D31"/></a:solidFill><a:prstDash val="dash"/></a:ln>"#), "{out}");
    }

    #[test]
    fn a_line_turned_off_takes_nothing_from_the_style() {
        // alpha_fill probe: an explicit `<a:ln><a:noFill/>` under lnRef idx=1
        // must stay without a width, or the parser strokes it.
        let xml = r#"<p:sp><p:spPr><a:solidFill><a:srgbClr val="000000"><a:alpha val="0"/></a:srgbClr></a:solidFill><a:ln><a:noFill/></a:ln></p:spPr><p:style><a:lnRef idx="1"><a:schemeClr val="accent1"/></a:lnRef><a:fillRef idx="3"><a:schemeClr val="accent1"/></a:fillRef></p:style></p:sp>"#;
        let out = norm(xml);
        assert!(out.contains("<a:ln><a:noFill/></a:ln>"), "{out}");
    }

    #[test]
    fn lum_mod_and_clr_map() {
        // bg1 lumMod 50% = 808080 grey; tx1 lumMod 75% lumOff 25% = 404040
        let xml = r#"<a:p><a:solidFill><a:schemeClr val="bg1"><a:lumMod val="50000"/></a:schemeClr></a:solidFill><a:solidFill><a:schemeClr val="tx1"><a:lumMod val="75000"/><a:lumOff val="25000"/></a:schemeClr></a:solidFill><a:solidFill><a:schemeClr val="accent1"><a:lumMod val="20000"/><a:lumOff val="80000"/></a:schemeClr></a:solidFill></a:p>"#;
        let out = norm(xml);
        assert!(out.contains(r#"<a:srgbClr val="808080"/>"#), "{out}");
        assert!(out.contains(r#"<a:srgbClr val="404040"/>"#), "{out}");
        // PowerPoint's "Accent 1, Lighter 80%" of 4472C4 is DAE3F3.
        assert!(out.contains(r#"<a:srgbClr val="DAE3F3"/>"#), "{out}");
    }

    #[test]
    fn clr_map_routes_bg1_to_dark_slot() {
        let mut map = default_clr_map();
        map.insert("bg1".into(), "dk1".into());
        let out = normalize_part(r#"<a:solidFill><a:schemeClr val="bg1"/></a:solidFill>"#, &theme(), &map).unwrap();
        assert!(out.contains(r#"<a:srgbClr val="000000"/>"#), "{out}");
    }

    #[test]
    fn alpha_is_kept_and_plain_srgb_untouched() {
        let xml = r#"<x><a:srgbClr val="112233"><a:alpha val="50000"/></a:srgbClr><a:schemeClr val="accent1"><a:alpha val="25000"/></a:schemeClr></x>"#;
        let out = norm(xml);
        assert!(out.contains(r#"<a:srgbClr val="112233"><a:alpha val="50000"/></a:srgbClr>"#), "{out}");
        assert!(out.contains(r#"<a:srgbClr val="4472C4"><a:alpha val="25000"/></a:srgbClr>"#), "{out}");
    }

    #[test]
    fn bg_ref_becomes_bg_pr() {
        let xml = r#"<p:bg><p:bgRef idx="1001"><a:schemeClr val="bg1"/></p:bgRef></p:bg>"#;
        let out = norm(xml);
        assert_eq!(out, r#"<p:bg><p:bgPr><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill><a:effectLst/></p:bgPr></p:bg>"#);
    }

    #[test]
    fn font_tokens_resolve_and_empty_faces_drop() {
        let xml = r#"<a:rPr><a:latin typeface="+mj-lt"/><a:ea typeface="+mn-ea"/></a:rPr>"#;
        let out = norm(xml);
        assert_eq!(out, r#"<a:rPr><a:latin typeface="Calibri Light"/></a:rPr>"#);
    }

    #[test]
    fn default_table_style_paints_header_bands_and_borders() {
        let xml = r#"<a:tbl><a:tblPr firstRow="1" bandRow="1"><a:tableStyleId>{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}</a:tableStyleId></a:tblPr><a:tblGrid><a:gridCol w="1"/><a:gridCol w="1"/></a:tblGrid><a:tr h="1"><a:tc><a:txBody><a:bodyPr/><a:p><a:r><a:rPr lang="en-US"/><a:t>H</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc><a:tc><a:txBody><a:bodyPr/><a:p><a:r><a:t>H2</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc></a:tr><a:tr h="1"><a:tc><a:txBody><a:bodyPr/><a:p><a:r><a:t>a</a:t></a:r></a:p></a:txBody><a:tcPr><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:tcPr></a:tc><a:tc><a:txBody><a:bodyPr/><a:p><a:r><a:t>b</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc></a:tr><a:tr h="1"><a:tc><a:txBody><a:bodyPr/><a:p><a:r><a:t>c</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc><a:tc><a:txBody><a:bodyPr/><a:p><a:r><a:t>d</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc></a:tr></a:tbl>"#;
        let out = normalize_part_with_tables(xml, &theme(), &default_clr_map(), &HashMap::new()).unwrap();
        // header: accent1 fill, bold white text, 3pt white bottom rule
        assert!(out.contains(r#"<a:rPr lang="en-US" b="1"><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a:rPr><a:t>H</a:t>"#), "{out}");
        assert!(out.contains(r#"<a:lnB w="38100" cmpd="sng"><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a:lnB><a:solidFill><a:srgbClr val="4472C4"/></a:solidFill></a:tcPr>"#), "{out}");
        // first body row = band1H (accent1 tint 40% = CFD5EA, PowerPoint's own
        // banded-row colour), unless the cell has its own fill
        assert!(out.contains(r#"<a:srgbClr val="FF0000"/></a:solidFill></a:tcPr>"#), "{out}");
        assert!(out.contains(r#"<a:srgbClr val="CFD5EA"/></a:solidFill></a:tcPr></a:tc></a:tr><a:tr h="1">"#), "{out}");
        // second body row = band2H, i.e. the whole-table tint 20% = E9EBF5, dk1 text
        assert!(out.contains(r#"<a:srgbClr val="E9EBF5"/></a:solidFill></a:tcPr></a:tc></a:tr></a:tbl>"#), "{out}");
        assert!(out.contains(r#"<a:r><a:rPr><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:rPr><a:t>d</a:t>"#), "{out}");
    }

    #[test]
    fn text_and_entities_round_trip() {
        let xml = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n<p:sld xmlns:a=\"x\"><a:t xml:space=\"preserve\"> A &amp; B &lt;c&gt; </a:t><a:t>日本語</a:t></p:sld>";
        assert_eq!(norm(xml), xml);
    }
}
