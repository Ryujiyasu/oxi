// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! Charts, drawn from the values the chart part caches.
//!
//! A chart part (`word/charts/chartN.xml`) keeps, beside the formulas that
//! point into its embedded workbook, the last values Word showed
//! (`c:numCache` / `c:strCache`). Those are enough to draw it without the
//! workbook: this module lays the chart out -- title, legend, plot area, axes
//! -- and turns it into the shapes and text boxes of a group, the form a
//! SmartArt or a Word drawing group takes, so the layout draws it the same
//! way.
//!
//! Word's own chart layout is not public; the placement here follows its
//! visible defaults (title on top, legend at its stated side, axis labels
//! outside the plot, Excel's automatic axis scale) and the manual layout the
//! part records when the author moved the plot area.

use std::collections::HashMap;

use super::dml::{average, color_of, parse_tree, run, text_box, Group, Node};
use super::theme::ThemeColors;
use crate::ir::{Alignment, Block, Paragraph, ParagraphStyle, PathSeg, RunStyle, VectorShape};

/// Plot kinds this module draws. 3-D kinds are drawn flat.
const DRAWN: &[&str] = &[
    "barChart", "bar3DChart", "lineChart", "line3DChart", "areaChart", "area3DChart",
    "pieChart", "pie3DChart", "doughnutChart", "scatterChart",
];

const TEXT_GREY: &str = "595959";
const GRID_GREY: &str = "D9D9D9";

/// A shape-properties fill: solid, or a gradient as the mean of its stops.
fn area_fill(sp: &Node, xml: &str, theme: &ThemeColors) -> Option<String> {
    if let Some(f) = sp.child("solidFill") {
        return color_of(f, xml, theme);
    }
    let stops = sp.path(&["gradFill", "gsLst"])?;
    average(stops.children.iter().filter_map(|gs| color_of(gs, xml, theme)))
}

fn plots(plot_area: &Node) -> impl Iterator<Item = &Node> {
    plot_area.children.iter().filter(|c| c.name.ends_with("Chart"))
}

/// Whether every plot of the chart part is one this module draws.
pub(super) fn is_drawable(xml: &str) -> bool {
    let Some(root) = parse_tree(xml) else { return false };
    let Some(area) = root.path(&["chart", "plotArea"]) else { return false };
    let mut any = false;
    for p in plots(area) {
        if !DRAWN.contains(&p.name.as_str()) {
            return false;
        }
        any = true;
    }
    any
}

// ---------- data ----------

struct Series {
    name: String,
    cats: Vec<String>,
    vals: Vec<Option<f64>>,
    xs: Vec<Option<f64>>,
    format: String,
    fill: Option<String>,
    line: Option<String>,
    line_width: Option<f32>,
    line_none: bool,
    marker: bool,
    point_fills: HashMap<usize, String>,
    /// A cached point's own `formatCode` (c:pt/@formatCode), over `format`.
    point_formats: HashMap<usize, String>,
}

impl Series {
    /// OXI_CHART_POINT_FORMAT (2026-11-07, default ON, opt-out _DISABLE): a value
    /// label takes its cached point's own formatCode. blind-J ja/reports__7a8daf36
    /// chart2: the cache's formatCode is 0.0% with points 0/2/3 marked 0%; Word's
    /// labels read 64% / 47.6% / 77% / 78%, Oxi gave 63.6% / 47.6% / 77.3% / 77.8%.
    fn format_at(&self, i: usize) -> &str {
        if !crate::env_flag!("OXI_CHART_POINT_FORMAT_DISABLE") {
            if let Some(f) = self.point_formats.get(&i) {
                return f;
            }
        }
        if self.format.is_empty() { "General" } else { &self.format }
    }
}

/// The per-point formatCode overrides of a value cache.
fn point_formats(n: Option<&Node>) -> HashMap<usize, String> {
    let Some(c) = n
        .and_then(|n| n.child("numRef"))
        .and_then(|r| r.child("numCache"))
        .or_else(|| n.and_then(|n| n.child("numLit")))
    else {
        return HashMap::new();
    };
    c.children
        .iter()
        .filter(|p| p.name == "pt")
        .filter_map(|p| Some((p.num("idx")? as usize, p.attr("formatCode")?.to_string())))
        .collect()
}

/// The cached points of a reference or literal: (values by index, format).
fn cache(n: Option<&Node>) -> (Vec<Option<String>>, String) {
    let Some(n) = n else { return (Vec::new(), String::new()) };
    let c = ["strRef", "numRef", "multiLvlStrRef"]
        .iter()
        .find_map(|r| n.child(r))
        .and_then(|r| r.child("strCache").or_else(|| r.child("numCache")).or_else(|| r.child("multiLvlStrCache")))
        .or_else(|| n.child("strLit"))
        .or_else(|| n.child("numLit"));
    let Some(c) = c else { return (Vec::new(), String::new()) };
    // A multi-level category cache: the innermost level labels the points.
    let c = c.child("lvl").unwrap_or(c);
    let count = c.child("ptCount").and_then(|p| p.num("val")).unwrap_or(0.0) as usize;
    let mut out = vec![None; count];
    for pt in c.children.iter().filter(|p| p.name == "pt") {
        let Some(idx) = pt.num("idx").map(|i| i as usize) else { continue };
        if idx >= out.len() {
            out.resize(idx + 1, None);
        }
        out[idx] = pt.child("v").map(|v| v.text.clone());
    }
    let format = c.child("formatCode").map(|f| f.text.clone()).unwrap_or_default();
    (out, format)
}

fn numbers(v: Vec<Option<String>>) -> Vec<Option<f64>> {
    v.into_iter().map(|s| s.and_then(|s| s.trim().parse().ok())).collect()
}

fn series(ser: &Node, xml: &str, theme: &ThemeColors) -> Series {
    let (name, _) = cache(ser.child("tx"));
    let name = name.into_iter().flatten().next().or_else(|| ser.path(&["tx", "v"]).map(|v| v.text.clone())).unwrap_or_default();
    let (cats, cat_format) = cache(ser.child("cat").or_else(|| ser.child("xVal")));
    let (vals, format) = cache(ser.child("val").or_else(|| ser.child("yVal")));
    let (xs, _) = cache(ser.child("xVal"));
    let sp = ser.child("spPr");
    let fill = sp.and_then(|s| s.child("solidFill")).and_then(|f| color_of(f, xml, theme))
        .or_else(|| sp.and_then(|s| s.path(&["gradFill", "gsLst", "gs"])).and_then(|f| color_of(f, xml, theme)));
    let ln = sp.and_then(|s| s.child("ln"));
    let point_fills = ser
        .children
        .iter()
        .filter(|c| c.name == "dPt")
        .filter_map(|d| Some((d.path(&["idx"])?.num("val")? as usize, color_of(d.path(&["spPr", "solidFill"])?, xml, theme)?)))
        .collect();
    let cat_numbers = numbers(cats.clone());
    Series {
        name,
        cats: cats
            .into_iter()
            .zip(cat_numbers)
            .map(|(c, n)| match n {
                Some(v) if !cat_format.is_empty() && cat_format != "General" => format_number(v, &cat_format),
                _ => c.unwrap_or_default(),
            })
            .collect(),
        vals: numbers(vals),
        xs: numbers(xs),
        format,
        fill,
        line: ln.and_then(|l| l.child("solidFill")).and_then(|f| color_of(f, xml, theme)),
        line_width: ln.and_then(|l| l.num("w")).map(|w| w / 12700.0),
        line_none: ln.is_some_and(|l| l.child("noFill").is_some()),
        marker: ser.path(&["marker", "symbol"]).and_then(|s| s.attr("val")).is_some_and(|v| v != "none"),
        point_fills,
        point_formats: point_formats(ser.child("val").or_else(|| ser.child("yVal"))),
    }
}

/// The colour Office gives series `idx` when the part names none: the six
/// accents, then darker, then lighter rounds of them.
fn default_color(idx: usize, theme: &ThemeColors) -> String {
    let base = theme.resolve(&format!("accent{}", idx % 6 + 1)).cloned().unwrap_or_else(|| "4472C4".into());
    let round = idx / 6;
    if round == 0 {
        return base;
    }
    let Ok(v) = u32::from_str_radix(base.get(..6).unwrap_or("4472C4"), 16) else { return base };
    let ch = [(v >> 16) as f32, ((v >> 8) & 0xFF) as f32, (v & 0xFF) as f32];
    let ch = if round % 2 == 1 { ch.map(|c| c * 0.6) } else { ch.map(|c| c + (255.0 - c) * 0.4) };
    format!("{:02X}{:02X}{:02X}", ch[0] as u8, ch[1] as u8, ch[2] as u8)
}

// ---------- numbers ----------

/// `v` in an Excel number format, for the common codes: General, fixed
/// decimals, thousands separators, percent, and literal prefixes/suffixes.
fn format_number(v: f64, code: &str) -> String {
    let code = code.split(';').next().unwrap_or("").trim();
    if code.is_empty() || code.eq_ignore_ascii_case("general") {
        return general(v);
    }
    let percent = code.contains('%');
    let v = if percent { v * 100.0 } else { v };
    let decimals = code.split('.').nth(1).map_or(0, |d| d.chars().take_while(|c| *c == '0' || *c == '#').count());
    let thousands = code.contains(",#") || code.contains("#,");
    // Excel rounds halves away from zero; `{:.N}` rounds them to even.
    let unit = 10f64.powi(decimals as i32);
    let mut body = format!("{:.*}", decimals, (v.abs() * unit).round() / unit);
    if thousands {
        let (int, frac) = body.split_once('.').map_or((body.clone(), None), |(i, f)| (i.to_string(), Some(f.to_string())));
        let mut grouped = String::new();
        for (i, ch) in int.chars().enumerate() {
            if i > 0 && (int.len() - i) % 3 == 0 {
                grouped.push(',');
            }
            grouped.push(ch);
        }
        body = match frac {
            Some(f) => format!("{grouped}.{f}"),
            None => grouped,
        };
    }
    // Literal text around the digits: "$", "¥", "円", quoted strings.
    let first = code.find(['0', '#', '?']).unwrap_or(code.len());
    let last = code.rfind(['0', '#', '?']).map_or(code.len(), |i| i + 1);
    let lit = |s: &str| s.replace(['"', '\\', '%'], "").replace('_', "").trim().to_string();
    let sign = if v < 0.0 && body.chars().any(|c| c.is_ascii_digit() && c != '0') { "-" } else { "" };
    format!("{sign}{}{body}{}{}", lit(&code[..first]), lit(&code[last..]), if percent { "%" } else { "" })
}

fn general(v: f64) -> String {
    if v == v.trunc() && v.abs() < 1e15 {
        return format!("{}", v as i64);
    }
    let s = format!("{:.10}", v);
    s.trim_end_matches('0').trim_end_matches('.').to_string()
}

/// Excel's automatic axis: from zero when the data allow, a 1-2-5 major unit
/// for about five intervals, and the top a unit above the data.
fn scale(lo: f64, hi: f64, axis: Option<&Node>) -> (f64, f64, f64) {
    let fixed = |name: &str| axis.and_then(|a| a.path(&["scaling", name])).and_then(|n| n.num("val")).map(f64::from);
    let major_fixed = axis.and_then(|a| a.child("majorUnit")).and_then(|n| n.num("val")).map(f64::from);
    let (lo, hi) = if lo.is_finite() && hi.is_finite() { (lo, hi) } else { (0.0, 1.0) };
    let mut min = fixed("min").unwrap_or(if lo >= 0.0 && lo <= hi * 5.0 / 6.0 { 0.0 } else { lo });
    let mut max = fixed("max").unwrap_or(if hi <= 0.0 { 0.0 } else { hi });
    if max <= min {
        max = min + 1.0;
    }
    let major = major_fixed.unwrap_or_else(|| major_unit(max - min));
    if fixed("min").is_none() && min != 0.0 {
        min = (min / major).floor() * major;
    }
    if fixed("max").is_none() {
        max = ((hi + 0.05 * (hi - min)) / major).ceil() * major;
        if max <= min {
            max = min + major;
        }
    }
    (min, max, major)
}

/// The 1-2-5 major unit that cuts `range` into the count of intervals nearest
/// six, the larger unit on a tie: 0..4.5 by 1, -3..8 by 2.
fn major_unit(range: f64) -> f64 {
    if range <= 0.0 || !range.is_finite() {
        return 1.0;
    }
    let mag = 10f64.powf((range / 6.0).log10().floor());
    [0.5, 1.0, 2.0, 5.0, 10.0]
        .iter()
        .map(|m| m * mag)
        .min_by(|a, b| ((range / a - 6.0).abs(), -a).partial_cmp(&((range / b - 6.0).abs(), -b)).unwrap())
        .unwrap_or(mag)
}

fn nice_step(raw: f64) -> f64 {
    if raw <= 0.0 || !raw.is_finite() {
        return 1.0;
    }
    let mag = 10f64.powf(raw.log10().floor());
    let n = raw / mag;
    mag * if n <= 1.0 { 1.0 } else if n <= 2.0 { 2.0 } else if n <= 5.0 { 5.0 } else { 10.0 }
}

// ---------- text ----------

#[derive(Clone)]
struct TextStyle {
    size: f32,
    bold: bool,
    color: String,
    latin: Option<String>,
    ea: Option<String>,
}

/// The text style an element's `txPr` sets, each property it leaves out
/// taken from `inherit` (the chart's own `txPr`, then Office's defaults).
fn text_style(n: Option<&Node>, xml: &str, theme: &ThemeColors, inherit: &TextStyle) -> TextStyle {
    let d = n.and_then(|n| n.path(&["txPr", "p", "pPr", "defRPr"]));
    apply_rpr(d, xml, theme, inherit.clone())
}

/// `st` with the size, weight and colour a run-properties element sets.
fn apply_rpr(d: Option<&Node>, xml: &str, theme: &ThemeColors, mut st: TextStyle) -> TextStyle {
    if let Some(d) = d {
        if let Some(sz) = d.num("sz") {
            st.size = sz / 100.0;
        }
        if let Some(b) = d.attr("b") {
            st.bold = b == "1";
        }
        if let Some(c) = d.child("solidFill").and_then(|f| color_of(f, xml, theme)) {
            st.color = c;
        }
    }
    st
}

/// A rough advance for laying labels out: full width for East Asian text,
/// a little over half for Latin.
fn text_width(text: &str, size: f32) -> f32 {
    text.chars().map(|c| if (c as u32) >= 0x2E80 { size } else { size * 0.56 }).sum()
}

fn label(g: &mut Group, text: &str, st: &TextStyle, rect: (f32, f32, f32, f32), align: Alignment) {
    if text.is_empty() || rect.2 <= 0.0 || rect.3 <= 0.0 {
        return;
    }
    let runs = text
        .split('\n')
        .enumerate()
        .map(|(i, line)| {
            run(
                if i == 0 { line.to_string() } else { format!("\n{line}") },
                RunStyle {
                    font_size: Some(st.size),
                    bold: st.bold,
                    color: Some(st.color.clone()),
                    font_family: st.latin.clone(),
                    font_family_east_asia: st.ea.clone(),
                    ..RunStyle::default()
                },
            )
        })
        .collect();
    let para = Block::Paragraph(Paragraph {
        runs,
        style: ParagraphStyle {
            space_before: Some(0.0),
            space_after: Some(0.0),
            snap_to_grid: false,
            ..ParagraphStyle::default()
        },
        alignment: align,
        shapes: Vec::new(),
        ppr_change: None,
        paragraph_mark_revision: None,
    });
    g.texts.push(text_box(vec![para], [0.0; 4], "middle", false, rect));
}

fn rich_text(n: &Node) -> Option<String> {
    let rich = n.path(&["tx", "rich"])?;
    let lines: Vec<String> = rich
        .children
        .iter()
        .filter(|p| p.name == "p")
        .map(|p| p.children.iter().filter(|r| r.name == "r" || r.name == "fld").filter_map(|r| r.child("t")).map(|t| t.text.as_str()).collect())
        .collect();
    let text = lines.join("\n");
    (!text.trim().is_empty()).then_some(text)
}

// ---------- shapes ----------

fn rect(g: &mut Group, (x, y, w, h): (f32, f32, f32, f32), fill: Option<String>, stroke: Option<(String, f32)>) {
    if w <= 0.0 || h <= 0.0 {
        return;
    }
    let (stroke, stroke_width) = stroke.map_or((None, 0.0), |(c, w)| (Some(c), w));
    g.shapes.push(VectorShape { x, y, w, h, fill, stroke, stroke_width, is_line: false, path: Vec::new(), image: None });
}

fn line(g: &mut Group, (x1, y1): (f32, f32), (x2, y2): (f32, f32), color: &str, width: f32) {
    g.shapes.push(VectorShape {
        x: x1,
        y: y1,
        w: x2 - x1,
        h: y2 - y1,
        fill: None,
        stroke: Some(color.to_string()),
        stroke_width: width,
        is_line: true,
        path: Vec::new(),
        image: None,
    });
}

/// A path in absolute points, normalized over its own bounding box.
fn path(g: &mut Group, segs: Vec<PathSeg>, fill: Option<String>, stroke: Option<(String, f32)>) {
    let pts: Vec<(f32, f32)> = segs
        .iter()
        .flat_map(|s| match *s {
            PathSeg::M(x, y) | PathSeg::L(x, y) => vec![(x, y)],
            PathSeg::C(a, b, c, d, e, f) => vec![(a, b), (c, d), (e, f)],
            PathSeg::Z => vec![],
        })
        .collect();
    if pts.is_empty() {
        return;
    }
    let (x0, y0) = pts.iter().fold((f32::MAX, f32::MAX), |(a, b), p| (a.min(p.0), b.min(p.1)));
    let (x1, y1) = pts.iter().fold((f32::MIN, f32::MIN), |(a, b), p| (a.max(p.0), b.max(p.1)));
    let (w, h) = ((x1 - x0).max(0.01), (y1 - y0).max(0.01));
    let n = |x: f32, y: f32| ((x - x0) / w, (y - y0) / h);
    let segs = segs
        .into_iter()
        .map(|s| match s {
            PathSeg::M(x, y) => { let p = n(x, y); PathSeg::M(p.0, p.1) }
            PathSeg::L(x, y) => { let p = n(x, y); PathSeg::L(p.0, p.1) }
            PathSeg::C(a, b, c, d, e, f) => { let (p, q, r) = (n(a, b), n(c, d), n(e, f)); PathSeg::C(p.0, p.1, q.0, q.1, r.0, r.1) }
            PathSeg::Z => PathSeg::Z,
        })
        .collect();
    let (stroke, stroke_width) = stroke.map_or((None, 0.0), |(c, w)| (Some(c), w));
    g.shapes.push(VectorShape { x: x0, y: y0, w, h, fill, stroke, stroke_width, is_line: false, path: segs, image: None });
}

/// A point on a circle at `deg` clockwise from twelve o'clock.
fn on_circle(cx: f32, cy: f32, r: f32, deg: f32) -> (f32, f32) {
    let a = (deg - 90.0).to_radians();
    (cx + r * a.cos(), cy + r * a.sin())
}

/// Cubic segments along a circle from `from` to `to` degrees (clockwise
/// from twelve o'clock; `to` < `from` runs back).
fn arc(out: &mut Vec<PathSeg>, cx: f32, cy: f32, r: f32, from: f32, to: f32) {
    let n = ((to - from).abs() / 90.0).ceil().max(1.0) as usize;
    let step = (to - from) / n as f32;
    let k = 4.0 / 3.0 * (step.to_radians() / 4.0).tan();
    for i in 0..n {
        let (a0, a1) = (from + step * i as f32, from + step * (i + 1) as f32);
        let (t0, t1) = ((a0 - 90.0).to_radians(), (a1 - 90.0).to_radians());
        let p0 = on_circle(cx, cy, r, a0);
        let p3 = on_circle(cx, cy, r, a1);
        out.push(PathSeg::C(
            p0.0 - k * r * t0.sin(), p0.1 + k * r * t0.cos(),
            p3.0 + k * r * t1.sin(), p3.1 - k * r * t1.cos(),
            p3.0, p3.1,
        ));
    }
}

// ---------- layout ----------

fn legend_entries(kind: &str, sers: &[Series], vary: bool, theme: &ThemeColors) -> Vec<(String, String)> {
    if vary && sers.len() == 1 {
        let s = &sers[0];
        return s.cats.iter().enumerate().map(|(i, c)| (c.clone(), s.point_fills.get(&i).cloned().unwrap_or_else(|| default_color(i, theme)))).collect();
    }
    sers.iter()
        .enumerate()
        .map(|(i, s)| {
            let c = if kind.starts_with("line") || kind == "scatterChart" { s.line.clone().or_else(|| s.fill.clone()) } else { s.fill.clone() };
            (s.name.clone(), c.unwrap_or_else(|| default_color(i, theme)))
        })
        .collect()
}

/// The chart drawn by `xml`, a chart part, in a frame of `w` × `h` points.
pub(super) fn from_chart(xml: &str, theme: &ThemeColors, w: f32, h: f32) -> Option<Group> {
    if w <= 0.0 || h <= 0.0 || !is_drawable(xml) {
        return None;
    }
    let root = parse_tree(xml)?;
    let chart = root.child("chart")?;
    let area = chart.child("plotArea")?;
    let mut g = Group { shapes: Vec::new(), texts: Vec::new() };

    // Chart area.
    if let Some(sp) = root.child("spPr") {
        let fill = area_fill(sp, xml, theme);
        let ln = sp.child("ln").and_then(|l| Some((color_of(l.child("solidFill")?, xml, theme)?, l.num("w").map_or(0.75, |w| w / 12700.0))));
        rect(&mut g, (0.0, 0.0, w, h), fill, ln);
    }
    let office = TextStyle {
        size: 9.0,
        bold: false,
        color: TEXT_GREY.into(),
        latin: theme.minor_font.clone(),
        ea: theme.minor_font_ea.clone(),
    };
    let base = text_style(Some(&root), xml, theme, &office);

    let plot_nodes: Vec<&Node> = plots(area).collect();
    let first = plot_nodes.first()?;
    let kind = first.name.as_str();
    let all_series: Vec<Series> = plot_nodes.iter().flat_map(|p| p.children.iter().filter(|c| c.name == "ser").map(|s| series(s, xml, theme))).collect();
    let vary = first.child("varyColors").and_then(|v| v.attr("val")).map_or(kind.contains("pie") || kind.contains("doughnut"), |v| v == "1");

    let pad = 7.0;
    let (mut left, mut top, mut right, mut bottom) = (pad, pad, w - pad, h - pad);

    // Title: its own text, or the one series' name when there is none.
    let deleted = chart.child("autoTitleDeleted").and_then(|n| n.attr("val")) == Some("1");
    if let Some(t) = chart.child("title") {
        if !deleted || rich_text(t).is_some() {
            let text = rich_text(t).unwrap_or_else(|| if all_series.len() == 1 { all_series[0].name.clone() } else { "Chart Title".into() });
            // Office sets a title at 14pt; its own txPr, then its text's run
            // properties, override.
            let st = text_style(Some(t), xml, theme, &TextStyle { size: 14.0, ..base.clone() });
            let para = t.path(&["tx", "rich", "p"]);
            let st = apply_rpr(para.and_then(|p| p.path(&["pPr", "defRPr"])), xml, theme, st);
            let first_run = para.and_then(|p| p.children.iter().find(|r| r.name == "r")).and_then(|r| r.child("rPr"));
            let st = apply_rpr(first_run, xml, theme, st);
            let th = st.size * 1.3 * text.lines().count().max(1) as f32;
            label(&mut g, &text, &st, (left, top, right - left, th), Alignment::Center);
            top += th + 4.0;
        }
    }

    // Legend.
    if let Some(lg) = chart.child("legend") {
        let st = text_style(Some(lg), xml, theme, &base);
        let hidden: Vec<usize> = lg
            .children
            .iter()
            .filter(|e| e.name == "legendEntry" && e.path(&["delete"]).and_then(|d| d.attr("val")) == Some("1"))
            .filter_map(|e| e.path(&["idx"]).and_then(|i| i.num("val")).map(|i| i as usize))
            .collect();
        let entries: Vec<(String, String)> = legend_entries(kind, &all_series, vary, theme)
            .into_iter()
            .enumerate()
            .filter(|(i, _)| !hidden.contains(i))
            .map(|(_, e)| e)
            .collect();
        // A legend the author moved keeps its box and takes no room from the plot.
        let lm = lg.path(&["layout", "manualLayout"]);
        let lfrac = |k: &str| lm.and_then(|m| m.child(k)).and_then(|n| n.num("val"));
        let moved = match (lfrac("x"), lfrac("y"), lfrac("w"), lfrac("h")) {
            (Some(x), Some(y), Some(fw), Some(fh)) => Some((x * w, y * h, fw * w, fh * h)),
            _ => None,
        };
        let swatch = st.size * 0.6;
        let pos = lg.child("legendPos").and_then(|p| p.attr("val")).unwrap_or("r");
        let row_h = st.size * 1.4;
        let entry_w = |t: &str| swatch + 4.0 + text_width(t, st.size) + 2.0;
        let horizontal_row = matches!(pos, "b" | "t");
        match (pos, moved) {
            (_, Some((mx, my, mw, mh))) if horizontal_row => {
                let total: f32 = entries.iter().map(|(t, _)| entry_w(t) + 8.0).sum();
                let y = my + (mh - row_h) / 2.0;
                let mut x = mx + (mw - total).max(0.0) / 2.0;
                for (t, c) in &entries {
                    rect(&mut g, (x, y + (row_h - swatch) / 2.0, swatch, swatch), Some(c.clone()), None);
                    label(&mut g, t, &st, (x + swatch + 4.0, y, entry_w(t), row_h), Alignment::Left);
                    x += entry_w(t) + 8.0;
                }
            }
            (_, Some((mx, my, mw, mh))) => {
                let mut y = my + (mh - row_h * entries.len() as f32).max(0.0) / 2.0;
                for (t, c) in &entries {
                    rect(&mut g, (mx, y + (row_h - swatch) / 2.0, swatch, swatch), Some(c.clone()), None);
                    label(&mut g, t, &st, (mx + swatch + 4.0, y, (mw - swatch - 4.0).max(entry_w(t)), row_h), Alignment::Left);
                    y += row_h;
                }
            }
            ("b" | "t", None) => {
                let total: f32 = entries.iter().map(|(t, _)| entry_w(t) + 8.0).sum();
                let y = if pos == "b" { bottom - row_h } else { top };
                let mut x = left + ((right - left) - total).max(0.0) / 2.0;
                for (t, c) in &entries {
                    rect(&mut g, (x, y + (row_h - swatch) / 2.0, swatch, swatch), Some(c.clone()), None);
                    label(&mut g, t, &st, (x + swatch + 4.0, y, entry_w(t), row_h), Alignment::Left);
                    x += entry_w(t) + 8.0;
                }
                if pos == "b" { bottom -= row_h + 4.0 } else { top += row_h + 4.0 }
            }
            _ => {
                let col_w = entries.iter().map(|(t, _)| entry_w(t)).fold(0.0, f32::max).min((right - left) * 0.4);
                let x = if pos == "l" { left } else { right - col_w };
                let total = row_h * entries.len() as f32;
                let mut y = if pos == "tr" { top } else { top + ((bottom - top) - total).max(0.0) / 2.0 };
                for (t, c) in &entries {
                    rect(&mut g, (x, y + (row_h - swatch) / 2.0, swatch, swatch), Some(c.clone()), None);
                    label(&mut g, t, &st, (x + swatch + 4.0, y, col_w - swatch - 4.0, row_h), Alignment::Left);
                    y += row_h;
                }
                if pos == "l" { left += col_w + 6.0 } else { right -= col_w + 6.0 }
            }
        }
    }

    // Plot area: the recorded manual layout, else what is left.
    let manual = area.path(&["layout", "manualLayout"]);
    let frac = |k: &str| manual.and_then(|m| m.child(k)).and_then(|n| n.num("val"));
    let inner_manual = manual.and_then(|m| m.child("layoutTarget")).and_then(|t| t.attr("val")) == Some("inner");
    let plot = match (frac("x"), frac("y"), frac("w"), frac("h")) {
        (Some(x), Some(y), Some(fw), Some(fh)) => (x * w, y * h, fw * w, fh * h),
        _ => (left, top, right - left, bottom - top),
    };
    if let Some(sp) = area.child("spPr") {
        if let Some(fill) = area_fill(sp, xml, theme) {
            rect(&mut g, plot, Some(fill), None);
        }
    }

    match kind {
        "pieChart" | "pie3DChart" | "doughnutChart" => draw_pie(&mut g, first, &all_series, vary, plot, &base, xml, theme),
        "scatterChart" => draw_scatter(&mut g, area, &all_series, plot, inner_manual, &base, xml, theme),
        _ => draw_category(&mut g, area, &plot_nodes, &all_series, plot, inner_manual, &base, xml, theme),
    }
    Some(g)
}

fn axis<'a>(area: &'a Node, name: &str) -> Option<&'a Node> {
    area.children.iter().find(|c| c.name == name)
}

fn axis_hidden(ax: Option<&Node>) -> bool {
    ax.is_none_or(|a| a.child("delete").and_then(|d| d.attr("val")) == Some("1"))
}

fn axis_format(ax: Option<&Node>, fallback: &str) -> String {
    let nf = ax.and_then(|a| a.child("numFmt"));
    match nf {
        Some(n) if n.attr("sourceLinked") != Some("1") => n.attr("formatCode").unwrap_or("General").to_string(),
        _ => if fallback.is_empty() { "General".into() } else { fallback.to_string() },
    }
}

fn grid_color(ax: Option<&Node>, xml: &str, theme: &ThemeColors) -> Option<(String, f32)> {
    let gl = ax?.child("majorGridlines")?;
    let ln = gl.path(&["spPr", "ln"]);
    if ln.is_some_and(|l| l.child("noFill").is_some()) {
        return None;
    }
    let c = ln.and_then(|l| l.child("solidFill")).and_then(|f| color_of(f, xml, theme)).unwrap_or_else(|| GRID_GREY.into());
    Some((c, ln.and_then(|l| l.num("w")).map_or(0.75, |w| w / 12700.0)))
}

fn axis_line(ax: Option<&Node>, xml: &str, theme: &ThemeColors, default: Option<&str>) -> Option<(String, f32)> {
    let ln = ax?.path(&["spPr", "ln"]);
    if ln.is_some_and(|l| l.child("noFill").is_some()) {
        return None;
    }
    let c = ln.and_then(|l| l.child("solidFill")).and_then(|f| color_of(f, xml, theme)).or_else(|| default.map(str::to_string))?;
    Some((c, ln.and_then(|l| l.num("w")).map_or(0.75, |w| w / 12700.0)))
}

fn show_values(plot: &Node, ser_node_has: bool) -> bool {
    ser_node_has || plot.path(&["dLbls", "showVal"]).and_then(|n| n.attr("val")) == Some("1")
}

#[allow(clippy::too_many_arguments)]
fn draw_category(
    g: &mut Group,
    area: &Node,
    plots: &[&Node],
    sers: &[Series],
    plot: (f32, f32, f32, f32),
    inner_manual: bool,
    base: &TextStyle,
    xml: &str,
    theme: &ThemeColors,
) {
    let first = plots[0];
    let horizontal = first.path(&["barDir"]).and_then(|b| b.attr("val")) == Some("bar");
    let grouping = first.child("grouping").and_then(|n| n.attr("val")).unwrap_or("clustered");
    let stacked = grouping == "stacked" || grouping == "percentStacked";
    let percent = grouping == "percentStacked";
    let n_cat = sers.iter().map(|s| s.vals.len().max(s.cats.len())).max().unwrap_or(0).max(1);
    let cats: Vec<String> = (0..n_cat)
        .map(|i| sers.iter().find_map(|s| s.cats.get(i).filter(|c| !c.is_empty()).cloned()).unwrap_or_else(|| (i + 1).to_string()))
        .collect();

    // The values as drawn: stacked sums, or shares of the category total.
    let totals: Vec<f64> = (0..n_cat).map(|i| sers.iter().filter_map(|s| s.vals.get(i).copied().flatten()).map(f64::abs).sum()).collect();
    let value = |s: &Series, i: usize| -> Option<f64> {
        let v = s.vals.get(i).copied().flatten()?;
        Some(if percent { if totals[i] > 0.0 { v / totals[i] } else { 0.0 } } else { v })
    };
    let (mut lo, mut hi) = (f64::MAX, f64::MIN);
    for i in 0..n_cat {
        if stacked {
            let (mut pos, mut neg) = (0.0, 0.0);
            for s in sers {
                let v = value(s, i).unwrap_or(0.0);
                if v >= 0.0 { pos += v } else { neg += v }
            }
            lo = lo.min(neg);
            hi = hi.max(pos);
        } else {
            for s in sers {
                if let Some(v) = value(s, i) {
                    lo = lo.min(v);
                    hi = hi.max(v);
                }
            }
        }
    }
    let val_ax = axis(area, "valAx");
    let cat_ax = axis(area, "catAx").or_else(|| axis(area, "dateAx"));
    let (min, max, major) = if percent { (0.0, 1.0, 0.1_f64.max(nice_step(0.2))) } else { scale(lo, hi, val_ax) };
    let fmt = if percent { "0%".to_string() } else { axis_format(val_ax, sers.first().map_or("", |s| s.format.as_str())) };
    let ticks: Vec<f64> = (0..=((max - min) / major).round() as i64).map(|k| min + k as f64 * major).collect();
    let tick_labels: Vec<String> = ticks.iter().map(|t| format_number(*t, &fmt)).collect();

    let st_val = text_style(val_ax, xml, theme, base);
    let st_cat = text_style(cat_ax, xml, theme, base);
    let val_label_w = if axis_hidden(val_ax) { 0.0 } else { tick_labels.iter().map(|t| text_width(t, st_val.size)).fold(0.0, f32::max) + 6.0 };
    let cat_label_w = if axis_hidden(cat_ax) { 0.0 } else { cats.iter().map(|t| text_width(t, st_cat.size)).fold(0.0, f32::max) + 6.0 };
    let line_h = |st: &TextStyle| st.size * 1.4;

    // Inner plot: the manual inner box, else the plot area less the labels.
    let (px, py, pw, ph) = plot;
    let inner = if inner_manual {
        plot
    } else if horizontal {
        let lw = cat_label_w.min(pw * 0.4);
        let bh = if axis_hidden(val_ax) { 0.0 } else { line_h(&st_val) };
        (px + lw, py, pw - lw, ph - bh)
    } else {
        let bh = if axis_hidden(cat_ax) { 0.0 } else { line_h(&st_cat) * cats.iter().map(|c| c.lines().count().max(1)).max().unwrap_or(1) as f32 };
        (px + val_label_w, py, pw - val_label_w, ph - bh)
    };
    let (ix, iy, iw, ih) = inner;
    if iw <= 0.0 || ih <= 0.0 {
        return;
    }
    let span = (max - min) as f32;
    // Position of value v along the value axis, in points.
    let at = |v: f64| -> f32 {
        let f = ((v - min) as f32 / span).clamp(0.0, 1.0);
        if horizontal { ix + f * iw } else { iy + ih - f * ih }
    };
    let base_v = 0.0_f64.clamp(min, max);

    // Gridlines and value labels.
    let grid = grid_color(val_ax, xml, theme);
    for (t, txt) in ticks.iter().zip(&tick_labels) {
        let p = at(*t);
        if let Some((c, lw)) = &grid {
            if horizontal { line(g, (p, iy), (p, iy + ih), c, *lw) } else { line(g, (ix, p), (ix + iw, p), c, *lw) }
        }
        if !axis_hidden(val_ax) {
            if horizontal {
                let tw = text_width(txt, st_val.size) + 6.0;
                label(g, txt, &st_val, (p - tw / 2.0, iy + ih + 2.0, tw, line_h(&st_val)), Alignment::Center);
            } else {
                label(g, txt, &st_val, (ix - val_label_w, p - line_h(&st_val) / 2.0, val_label_w - 3.0, line_h(&st_val)), Alignment::Right);
            }
        }
    }

    // Category labels and axis line.
    let band = if horizontal { ih / n_cat as f32 } else { iw / n_cat as f32 };
    let band_at = |i: usize| -> f32 { if horizontal { iy + ih - (i + 1) as f32 * band } else { ix + i as f32 * band } };
    if !axis_hidden(cat_ax) {
        for (i, c) in cats.iter().enumerate() {
            if horizontal {
                label(g, c, &st_cat, (ix - cat_label_w.min(pw * 0.4), band_at(i), cat_label_w.min(pw * 0.4) - 3.0, band), Alignment::Right);
            } else {
                let lines = cats.iter().map(|c| c.lines().count().max(1)).max().unwrap_or(1) as f32;
                label(g, c, &st_cat, (band_at(i), iy + ih + 2.0, band, line_h(&st_cat) * lines), Alignment::Center);
            }
        }
    }
    let axis_pos = at(base_v);
    if let Some((c, lw)) = axis_line(cat_ax, xml, theme, Some(GRID_GREY)) {
        if horizontal { line(g, (axis_pos, iy), (axis_pos, iy + ih), &c, lw) } else { line(g, (ix, axis_pos), (ix + iw, axis_pos), &c, lw) }
    }
    if let Some((c, lw)) = axis_line(val_ax, xml, theme, None) {
        if horizontal { line(g, (ix, iy + ih), (ix + iw, iy + ih), &c, lw) } else { line(g, (ix, iy), (ix, iy + ih), &c, lw) }
    }

    // The plots, in order (a combo chart draws its bars, then its lines).
    let mut si = 0;
    for p in plots {
        let n_ser = p.children.iter().filter(|c| c.name == "ser").count();
        let my = &sers[si..si + n_ser];
        let offset = si;
        si += n_ser;
        let pk = p.name.as_str();
        let labels = show_values(p, false);
        let st_lbl = text_style(p.child("dLbls"), xml, theme, base);
        if pk.starts_with("bar") {
            let gap = p.child("gapWidth").and_then(|n| n.num("val")).unwrap_or(150.0) / 100.0;
            let overlap = p.child("overlap").and_then(|n| n.num("val")).map_or(if stacked { 1.0 } else { 0.0 }, |o| o / 100.0);
            let group_w = band / (1.0 + gap);
            let n = if stacked { 1.0 } else { n_ser as f32 };
            let bw = group_w / (n - overlap * (n - 1.0)).max(0.01);
            for i in 0..n_cat {
                let (mut pos_acc, mut neg_acc) = (0.0_f64, 0.0_f64);
                for (j, s) in my.iter().enumerate() {
                    let Some(v) = value(s, i) else { continue };
                    let (from, to) = if stacked {
                        if v >= 0.0 { let f = pos_acc; pos_acc += v; (f, pos_acc) } else { let f = neg_acc; neg_acc += v; (f, neg_acc) }
                    } else {
                        (base_v, v)
                    };
                    let slot = if stacked { 0.0 } else { j as f32 * bw * (1.0 - overlap) };
                    // Horizontal groups stack upward, the first series lowest,
                    // as their categories do.
                    let start = if horizontal {
                        band_at(i) + (band + group_w) / 2.0 - bw - slot
                    } else {
                        band_at(i) + (band - group_w) / 2.0 + slot
                    };
                    let (a, b) = (at(from), at(to));
                    let fill = s.point_fills.get(&i).cloned().or_else(|| s.fill.clone()).unwrap_or_else(|| {
                        if p.child("varyColors").and_then(|v| v.attr("val")) == Some("1") && n_ser == 1 { default_color(i, theme) } else { default_color(offset + j, theme) }
                    });
                    let stroke = s.line.clone().filter(|_| !s.line_none).map(|c| (c, s.line_width.unwrap_or(0.75)));
                    let r = if horizontal { (a.min(b), start, (a - b).abs(), bw) } else { (start, a.min(b), bw, (a - b).abs()) };
                    rect(g, r, Some(fill), stroke);
                    if labels {
                        let txt = format_number(s.vals.get(i).copied().flatten().unwrap_or(0.0), s.format_at(i));
                        let lh = line_h(&st_lbl);
                        let lw = text_width(&txt, st_lbl.size) + 6.0;
                        let rr = if stacked {
                            (r.0 + r.2 / 2.0 - lw / 2.0, r.1 + r.3 / 2.0 - lh / 2.0, lw, lh)
                        } else if horizontal {
                            (r.0 + r.2 + 2.0, r.1 + r.3 / 2.0 - lh / 2.0, lw, lh)
                        } else {
                            (r.0 + r.2 / 2.0 - lw / 2.0, r.1 - lh, lw, lh)
                        };
                        label(g, &txt, &st_lbl, rr, Alignment::Center);
                    }
                }
            }
        } else {
            // line / area: one point per category, at the band's centre
            let area_kind = pk.starts_with("area");
            let p_stacked = matches!(p.child("grouping").and_then(|n| n.attr("val")), Some("stacked") | Some("percentStacked"));
            let mut acc = vec![0.0_f64; n_cat];
            for (j, s) in my.iter().enumerate() {
                let color = s.line.clone().or_else(|| s.fill.clone()).unwrap_or_else(|| default_color(offset + j, theme));
                let prev = acc.clone();
                let pts: Vec<Option<(f32, f32)>> = (0..n_cat)
                    .map(|i| {
                        let v = value(s, i)?;
                        let v = if p_stacked { acc[i] += v; acc[i] } else { v };
                        let x = if area_kind && n_cat > 1 { ix + iw * i as f32 / (n_cat - 1) as f32 } else { band_at(i) + band / 2.0 };
                        Some((x, at(v)))
                    })
                    .collect();
                if area_kind {
                    let present: Vec<(usize, (f32, f32))> = pts.iter().enumerate().filter_map(|(i, p)| p.map(|p| (i, p))).collect();
                    if present.len() >= 2 {
                        let mut segs = vec![PathSeg::M(present[0].1 .0, present[0].1 .1)];
                        segs.extend(present.iter().skip(1).map(|(_, (x, y))| PathSeg::L(*x, *y)));
                        for (i, (x, _)) in present.iter().rev() {
                            let floor = if p_stacked { at(prev[*i]) } else { at(base_v) };
                            segs.push(PathSeg::L(*x, floor));
                        }
                        segs.push(PathSeg::Z);
                        path(g, segs, Some(s.fill.clone().unwrap_or(color.clone())), None);
                    }
                } else {
                    if !s.line_none {
                        let mut segs = Vec::new();
                        let mut pen_down = false;
                        for p in &pts {
                            match p {
                                Some((x, y)) => { segs.push(if pen_down { PathSeg::L(*x, *y) } else { PathSeg::M(*x, *y) }); pen_down = true; }
                                None => pen_down = false,
                            }
                        }
                        if segs.len() >= 2 {
                            path(g, segs, None, Some((color.clone(), s.line_width.unwrap_or(2.25))));
                        }
                    }
                    if s.marker {
                        for (x, y) in pts.iter().flatten() {
                            marker(g, *x, *y, &color);
                        }
                    }
                }
                if labels {
                    for (i, p) in pts.iter().enumerate() {
                        let Some((x, y)) = p else { continue };
                        let txt = format_number(s.vals.get(i).copied().flatten().unwrap_or(0.0), s.format_at(i));
                        let lw = text_width(&txt, st_lbl.size) + 6.0;
                        label(g, &txt, &st_lbl, (x - lw / 2.0, y - line_h(&st_lbl) - 2.0, lw, line_h(&st_lbl)), Alignment::Center);
                    }
                }
            }
        }
    }
}

fn marker(g: &mut Group, x: f32, y: f32, color: &str) {
    let r = 2.5;
    let mut segs = vec![PathSeg::M(x, y - r)];
    arc(&mut segs, x, y, r, 0.0, 360.0);
    segs.push(PathSeg::Z);
    path(g, segs, Some(color.to_string()), None);
}

#[allow(clippy::too_many_arguments)]
fn draw_scatter(
    g: &mut Group,
    area: &Node,
    sers: &[Series],
    plot: (f32, f32, f32, f32),
    inner_manual: bool,
    base: &TextStyle,
    xml: &str,
    theme: &ThemeColors,
) {
    let axes: Vec<&Node> = area.children.iter().filter(|c| c.name == "valAx").collect();
    let x_ax = axes.iter().copied().find(|a| matches!(a.child("axPos").and_then(|p| p.attr("val")), Some("b") | Some("t")));
    let y_ax = axes.iter().copied().find(|a| matches!(a.child("axPos").and_then(|p| p.attr("val")), Some("l") | Some("r")));
    let xs_of = |s: &Series| -> Vec<Option<f64>> {
        if s.xs.iter().any(Option::is_some) { s.xs.clone() } else { (0..s.vals.len()).map(|i| Some(i as f64 + 1.0)).collect() }
    };
    let (mut xlo, mut xhi, mut ylo, mut yhi) = (f64::MAX, f64::MIN, f64::MAX, f64::MIN);
    for s in sers {
        for (x, y) in xs_of(s).iter().zip(&s.vals) {
            if let (Some(x), Some(y)) = (x, y) {
                xlo = xlo.min(*x); xhi = xhi.max(*x); ylo = ylo.min(*y); yhi = yhi.max(*y);
            }
        }
    }
    let (x0, x1, xm) = scale(xlo, xhi, x_ax);
    let (y0, y1, ym) = scale(ylo, yhi, y_ax);
    let st_x = text_style(x_ax, xml, theme, base);
    let st_y = text_style(y_ax, xml, theme, base);
    let yfmt = axis_format(y_ax, sers.first().map_or("", |s| s.format.as_str()));
    let xfmt = axis_format(x_ax, "General");
    let yticks: Vec<f64> = (0..=((y1 - y0) / ym).round() as i64).map(|k| y0 + k as f64 * ym).collect();
    let xticks: Vec<f64> = (0..=((x1 - x0) / xm).round() as i64).map(|k| x0 + k as f64 * xm).collect();
    let ylab_w = if axis_hidden(y_ax) { 0.0 } else { yticks.iter().map(|t| text_width(&format_number(*t, &yfmt), st_y.size)).fold(0.0, f32::max) + 6.0 };
    let xlab_h = if axis_hidden(x_ax) { 0.0 } else { st_x.size * 1.4 };
    let (px, py, pw, ph) = plot;
    let (ix, iy, iw, ih) = if inner_manual { plot } else { (px + ylab_w, py, pw - ylab_w, ph - xlab_h) };
    if iw <= 0.0 || ih <= 0.0 {
        return;
    }
    let fx = |v: f64| ix + ((v - x0) / (x1 - x0)) as f32 * iw;
    let fy = |v: f64| iy + ih - ((v - y0) / (y1 - y0)) as f32 * ih;
    let grid = grid_color(y_ax, xml, theme);
    for t in &yticks {
        if let Some((c, lw)) = &grid { line(g, (ix, fy(*t)), (ix + iw, fy(*t)), c, *lw); }
        if !axis_hidden(y_ax) {
            let h = st_y.size * 1.4;
            label(g, &format_number(*t, &yfmt), &st_y, (ix - ylab_w, fy(*t) - h / 2.0, ylab_w - 3.0, h), Alignment::Right);
        }
    }
    if let Some((c, lw)) = grid_color(x_ax, xml, theme) {
        for t in &xticks { line(g, (fx(*t), iy), (fx(*t), iy + ih), &c, lw); }
    }
    if !axis_hidden(x_ax) {
        for t in &xticks {
            let txt = format_number(*t, &xfmt);
            let tw = text_width(&txt, st_x.size) + 6.0;
            label(g, &txt, &st_x, (fx(*t) - tw / 2.0, iy + ih + 2.0, tw, xlab_h), Alignment::Center);
        }
    }
    if let Some((c, lw)) = axis_line(x_ax, xml, theme, Some(GRID_GREY)) {
        let y = fy(0.0_f64.clamp(y0, y1));
        line(g, (ix, y), (ix + iw, y), &c, lw);
    }
    for (j, s) in sers.iter().enumerate() {
        let color = s.line.clone().or_else(|| s.fill.clone()).unwrap_or_else(|| default_color(j, theme));
        let pts: Vec<(f32, f32)> = xs_of(s).iter().zip(&s.vals).filter_map(|(x, y)| Some((fx((*x)?), fy((*y)?)))).collect();
        if !s.line_none && pts.len() >= 2 {
            let mut segs = vec![PathSeg::M(pts[0].0, pts[0].1)];
            segs.extend(pts.iter().skip(1).map(|(x, y)| PathSeg::L(*x, *y)));
            path(g, segs, None, Some((color.clone(), s.line_width.unwrap_or(2.25))));
        }
        if s.marker || s.line_none {
            for (x, y) in &pts {
                marker(g, *x, *y, &color);
            }
        }
    }
}

#[allow(clippy::too_many_arguments)]
fn draw_pie(
    g: &mut Group,
    plot_node: &Node,
    sers: &[Series],
    vary: bool,
    plot: (f32, f32, f32, f32),
    base: &TextStyle,
    xml: &str,
    theme: &ThemeColors,
) {
    let Some(s) = sers.first() else { return };
    let (px, py, pw, ph) = plot;
    let r = pw.min(ph) / 2.0 * 0.95;
    let (cx, cy) = (px + pw / 2.0, py + ph / 2.0);
    let hole = if plot_node.name == "doughnutChart" {
        plot_node.child("holeSize").and_then(|n| n.num("val")).map_or(0.5, |v| v / 100.0) * r
    } else {
        0.0
    };
    let total: f64 = s.vals.iter().flatten().map(|v| v.abs()).sum();
    if total <= 0.0 {
        return;
    }
    let mut angle = plot_node.child("firstSliceAng").and_then(|n| n.num("val")).unwrap_or(0.0);
    let dl = plot_node.child("dLbls");
    let flag = |k: &str| dl.and_then(|d| d.child(k)).and_then(|n| n.attr("val")) == Some("1");
    let (show_val, show_pct, show_cat) = (flag("showVal"), flag("showPercent"), flag("showCatName"));
    let st_lbl = text_style(dl, xml, theme, base);
    for (i, v) in s.vals.iter().enumerate() {
        let Some(v) = v else { continue };
        let sweep = (v.abs() / total * 360.0) as f32;
        if sweep <= 0.0 {
            continue;
        }
        let fill = s.point_fills.get(&i).cloned().unwrap_or_else(|| if vary { default_color(i, theme) } else { s.fill.clone().unwrap_or_else(|| default_color(0, theme)) });
        let end = angle + sweep;
        let mut segs = Vec::new();
        if hole > 0.0 {
            let p = on_circle(cx, cy, r, angle);
            segs.push(PathSeg::M(p.0, p.1));
            arc(&mut segs, cx, cy, r, angle, end);
            let q = on_circle(cx, cy, hole, end);
            segs.push(PathSeg::L(q.0, q.1));
            arc(&mut segs, cx, cy, hole, end, angle);
        } else {
            segs.push(PathSeg::M(cx, cy));
            let p = on_circle(cx, cy, r, angle);
            segs.push(PathSeg::L(p.0, p.1));
            arc(&mut segs, cx, cy, r, angle, end);
        }
        segs.push(PathSeg::Z);
        path(g, segs, Some(fill), Some(("FFFFFF".into(), 0.75)));
        if show_val || show_pct || show_cat {
            let mut parts = Vec::new();
            if show_cat { parts.push(s.cats.get(i).cloned().unwrap_or_default()); }
            if show_val { parts.push(format_number(*v, s.format_at(i))); }
            if show_pct { parts.push(format!("{:.0}%", v.abs() / total * 100.0)); }
            let txt = parts.join("; ");
            let (lx, ly) = on_circle(cx, cy, (r + hole) / 2.0 + if hole > 0.0 { 0.0 } else { r * 0.2 }, angle + sweep / 2.0);
            let lw = text_width(&txt, st_lbl.size) + 6.0;
            let lh = st_lbl.size * 1.4;
            label(g, &txt, &st_lbl, (lx - lw / 2.0, ly - lh / 2.0, lw, lh), Alignment::Center);
        }
        angle = end;
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn theme() -> ThemeColors {
        let mut t = ThemeColors::default();
        for (k, v) in [("accent1", "4472C4"), ("accent2", "ED7D31"), ("accent3", "A5A5A5"), ("tx1", "000000"), ("dk1", "000000"), ("lt1", "FFFFFF")] {
            t.colors.insert(k.into(), v.into());
        }
        t
    }

    fn bar(dir: &str, grouping: &str) -> String {
        let ser = |i: usize, name: &str, vals: [f64; 3]| {
            let pts: String = vals.iter().enumerate().map(|(k, v)| format!(r#"<c:pt idx="{k}"><c:v>{v}</c:v></c:pt>"#)).collect();
            format!(r#"<c:ser><c:idx val="{i}"/><c:order val="{i}"/><c:tx><c:strRef><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>{name}</c:v></c:pt></c:strCache></c:strRef></c:tx>
<c:cat><c:strRef><c:strCache><c:ptCount val="3"/><c:pt idx="0"><c:v>A</c:v></c:pt><c:pt idx="1"><c:v>B</c:v></c:pt><c:pt idx="2"><c:v>C</c:v></c:pt></c:strCache></c:strRef></c:cat>
<c:val><c:numRef><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="3"/>{pts}</c:numCache></c:numRef></c:val></c:ser>"#)
        };
        format!(
            r#"<c:chartSpace xmlns:c="c" xmlns:a="a"><c:chart><c:autoTitleDeleted val="1"/><c:plotArea><c:layout/><c:barChart><c:barDir val="{dir}"/><c:grouping val="{grouping}"/><c:varyColors val="0"/>{}{}<c:gapWidth val="100"/><c:axId val="1"/><c:axId val="2"/></c:barChart>
<c:catAx><c:axId val="1"/><c:delete val="0"/><c:axPos val="b"/></c:catAx><c:valAx><c:axId val="2"/><c:delete val="0"/><c:axPos val="l"/><c:majorGridlines/></c:valAx></c:plotArea><c:legend><c:legendPos val="b"/></c:legend></c:chart></c:chartSpace>"#,
            ser(0, "North", [4.0, 2.0, 3.0]),
            ser(1, "South", [1.0, 3.0, 5.0]),
        )
    }

    #[test]
    fn drawable_kinds() {
        assert!(is_drawable(&bar("col", "clustered")));
        let radar = r#"<c:chartSpace xmlns:c="c"><c:chart><c:plotArea><c:radarChart/></c:plotArea></c:chart></c:chartSpace>"#;
        assert!(!is_drawable(radar));
    }

    #[test]
    fn clustered_columns_stand_on_the_axis_in_series_colours() {
        let g = from_chart(&bar("col", "clustered"), &theme(), 300.0, 200.0).expect("chart");
        let bars: Vec<&VectorShape> = g.shapes.iter().filter(|s| !s.is_line && s.path.is_empty() && s.h > 6.0).collect();
        assert_eq!(bars.len(), 6, "two series of three");
        // all bars share one baseline (value 0), and the tallest is South C (5)
        let base = bars[0].y + bars[0].h;
        assert!(bars.iter().all(|b| (b.y + b.h - base).abs() < 0.01));
        let tallest = bars.iter().max_by(|a, b| a.h.total_cmp(&b.h)).unwrap();
        assert_eq!(tallest.fill.as_deref(), Some("ED7D31"));
        // series 1 bars sit to the right of series 0 bars in each category
        assert!(bars[1].x > bars[0].x);
        // legend and labels are text boxes; the axis runs 0..6 by 1
        let texts: Vec<String> = g.texts.iter().filter_map(|t| match &t.blocks[0] { Block::Paragraph(p) => Some(p.runs.iter().map(|r| r.text.clone()).collect()), _ => None }).collect();
        for want in ["North", "South", "A", "B", "C", "0", "6"] {
            assert!(texts.iter().any(|t| t == want), "{want} in {texts:?}");
        }
    }

    #[test]
    fn horizontal_stacked_bars_add_up() {
        let g = from_chart(&bar("bar", "stacked"), &theme(), 300.0, 200.0).expect("chart");
        let bars: Vec<&VectorShape> = g.shapes.iter().filter(|s| !s.is_line && s.path.is_empty() && s.w > 1.0 && s.h > 6.0).collect();
        assert_eq!(bars.len(), 6);
        // category C: 3 + 5 -- the second segment starts where the first ends
        let c: Vec<&&VectorShape> = bars.iter().filter(|b| (b.y - bars[0].y).abs() > 0.01).collect();
        assert!(!c.is_empty());
        let mut ends: Vec<(f32, f32)> = bars.iter().map(|b| (b.y, b.x)).collect();
        ends.sort_by(|a, b| a.partial_cmp(b).unwrap());
        assert!(bars.iter().any(|a| bars.iter().any(|b| (a.y - b.y).abs() < 0.01 && (a.x + a.w - b.x).abs() < 0.01)));
    }

    /// In a horizontal bar chart the first series sits lowest in its group,
    /// as the first category sits lowest on the axis.
    #[test]
    fn horizontal_groups_put_the_first_series_lowest() {
        let g = from_chart(&bar("bar", "clustered"), &theme(), 300.0, 200.0).expect("chart");
        let north = g.shapes.iter().find(|s| s.fill.as_deref() == Some("4472C4") && s.path.is_empty() && !s.is_line && s.w > 6.0).unwrap();
        let south = g.shapes.iter().find(|s| s.fill.as_deref() == Some("ED7D31") && s.path.is_empty() && !s.is_line && s.w > 6.0).unwrap();
        assert!(north.y > south.y, "series 0 below series 1: {} vs {}", north.y, south.y);
    }

    #[test]
    fn axis_scale_and_formats() {
        assert_eq!(scale(0.0, 4.5, None), (0.0, 5.0, 1.0));
        assert_eq!(scale(-3.0, 8.0, None), (-4.0, 10.0, 2.0));
        assert_eq!(format_number(1234.5, "#,##0"), "1,235");
        assert_eq!(format_number(0.25, "0%"), "25%");
        assert_eq!(format_number(0.5, "General"), "0.5");
        assert_eq!(format_number(3.0, "General"), "3");
        assert_eq!(format_number(1500.0, "\"¥\"#,##0"), "¥1,500");
    }

    #[test]
    fn pie_slices_fill_the_circle() {
        let xml = r#"<c:chartSpace xmlns:c="c"><c:chart><c:plotArea><c:pieChart><c:varyColors val="1"/><c:ser><c:idx val="0"/>
<c:cat><c:strLit><c:ptCount val="2"/><c:pt idx="0"><c:v>x</c:v></c:pt><c:pt idx="1"><c:v>y</c:v></c:pt></c:strLit></c:cat>
<c:val><c:numLit><c:ptCount val="2"/><c:pt idx="0"><c:v>1</c:v></c:pt><c:pt idx="1"><c:v>3</c:v></c:pt></c:numLit></c:val></c:ser><c:firstSliceAng val="0"/></c:pieChart></c:plotArea></c:chart></c:chartSpace>"#;
        let g = from_chart(xml, &theme(), 200.0, 200.0).expect("pie");
        let slices: Vec<&VectorShape> = g.shapes.iter().filter(|s| !s.path.is_empty()).collect();
        assert_eq!(slices.len(), 2);
        assert_eq!(slices[0].fill.as_deref(), Some("4472C4"));
        assert_eq!(slices[1].fill.as_deref(), Some("ED7D31"));
        // the quarter slice spans the top-right quadrant: its box is about r x r
        assert!((slices[0].w - slices[0].h).abs() < 1.0);
    }
}
