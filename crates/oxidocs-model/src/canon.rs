// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! Compare two packages as Word would read them.
//!
//! Relationship ids become what they point at (an external URL, or a digest
//! of the target part's own canonical form), editing-session attributes and
//! layout caches are dropped, attribute order is ignored. Two packages that
//! compare equal here say the same thing.

use std::collections::hash_map::DefaultHasher;
use std::collections::HashMap;
use std::hash::{Hash, Hasher};

use crate::package::{Package, Rel, REL_ENDNOTES, REL_FOOTNOTES};
use crate::props::is_session_attr;
use crate::schema;
use crate::xml::{self, Element, Node};

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Diff {
    pub part: String,
    pub path: String,
    pub detail: String,
}

impl std::fmt::Display for Diff {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        write!(f, "{} {}: {}", self.part, self.path, self.detail)
    }
}

/// Layout caches and editing-session records (the settings' w:rsids, a
/// style's w:rsid).
const DROPPED: &[&str] = &["w:proofErr", "w:lastRenderedPageBreak", "w:rsids", "w:rsid"];

struct Canon<'a> {
    pkg: &'a Package,
    digests: HashMap<String, String>,
    /// Read neighbouring runs with the same properties as one run, and runs
    /// holding nothing as absent (see [`compare_content`]).
    runs_as_text: bool,
}

impl<'a> Canon<'a> {
    fn new(pkg: &'a Package) -> Self {
        Canon { pkg, digests: HashMap::new(), runs_as_text: false }
    }

    fn part(&mut self, name: &str) -> Option<Element> {
        let root = self.pkg.xml(name).ok()?;
        let rels: HashMap<String, Rel> = self.pkg.rels(name).into_iter().map(|r| (r.id.clone(), r)).collect();
        Some(self.element(&root, &rels))
    }

    fn element(&mut self, e: &Element, rels: &HashMap<String, Rel>) -> Element {
        self.element_in("", e, rels)
    }


    fn element_in(&mut self, parent: &str, e: &Element, rels: &HashMap<String, Rel>) -> Element {
        let normalized;
        let e = if e.name == "w:tr" && e.child("w:trHeight").is_some() {
            normalized = hoist_row_height(e);
            &normalized
        } else {
            e
        };
        let mut attrs: Vec<(String, String)> = e
            .attrs
            .iter()
            .filter(|(k, _)| {
                // Markup-compatibility processing hints and namespace
                // declarations say how to read the file, not what it says.
                !is_session_attr(k) && !k.starts_with("xmlns") && !k.starts_with("mc:") && k != "xml:space"
            })
            .map(|(k, v)| {
                if k.starts_with("r:") && !v.is_empty() {
                    (k.clone(), self.describe(rels.get(v)))
                } else {
                    (k.clone(), v.clone())
                }
            })
            .collect();
        attrs.sort();
        let mut children = Vec::new();
        for n in &e.children {
            match n {
                Node::Element(c) if DROPPED.contains(&c.name.as_str()) => {}
                Node::Element(c) => {
                    let ce = self.element_in(&e.name, c, rels);
                    // An empty property bag says nothing.
                    if is_bag(&e.name, &c.name) && ce.children.is_empty() && ce.attrs.is_empty() {
                        continue;
                    }
                    children.push(Node::Element(ce))
                }
                Node::Text(t) => children.push(Node::Text(t.clone())),
            }
        }
        // Adjacent text nodes merge (a dropped element may have split one).
        let mut merged: Vec<Node> = Vec::new();
        for n in children {
            match (merged.last_mut(), n) {
                (Some(Node::Text(a)), Node::Text(b)) => a.push_str(&b),
                (_, n) => merged.push(n),
            }
        }
        if self.runs_as_text {
            merged = join_runs(merged);
        }
        // Property bags are sets: their order means nothing, so both sides
        // are put in schema order.
        if let Some(order) = bag_order(parent, &e.name) {
            merged.sort_by_key(|n| match n {
                Node::Element(c) => crate::props::index(order, &c.name),
                Node::Text(_) => usize::MAX,
            });
        }
        Element { name: e.name.clone(), attrs, children: merged }
    }

    fn describe(&mut self, rel: Option<&Rel>) -> String {
        let Some(rel) = rel else { return "dangling".into() };
        let kind = rel.rel_type.rsplit('/').next().unwrap_or("").to_string();
        if rel.external {
            return format!("{kind}:ext:{}", rel.target);
        }
        if !self.pkg.files.contains_key(&rel.target) {
            return format!("{kind}:missing");
        }
        if let Some(d) = self.digests.get(&rel.target) {
            return format!("{kind}:{d}");
        }
        // Guard against cycles while the digest is being computed.
        self.digests.insert(rel.target.clone(), "cycle".into());
        let ct = self.pkg.content_type(&rel.target);
        let mut h = DefaultHasher::new();
        ct.hash(&mut h);
        let is_xml = ct.ends_with("+xml") || ct.ends_with("/xml") || rel.target.ends_with(".xml");
        match is_xml.then(|| self.part(&rel.target)).flatten() {
            Some(e) => xml::write(&e, false).hash(&mut h),
            None => self.pkg.files[&rel.target].hash(&mut h),
        }
        let d = format!("{:016x}", h.finish());
        self.digests.insert(rel.target.clone(), d.clone());
        format!("{kind}:{d}")
    }
}

fn is_bag(parent: &str, name: &str) -> bool {
    schema::container_order(parent, name).is_some() && matches!(name, "w:pPr" | "w:rPr" | "w:tblPr" | "w:trPr" | "w:tcPr" | "w:tblPrEx" | "w:sectPr")
}

/// The order a property container (or nested property group, or a list of
/// named settings) uses. Children of these are sets: order means nothing.
fn bag_order(parent: &str, name: &str) -> Option<&'static [&'static str]> {
    use crate::docparts::{ABSTRACT_META, FONT, LEVEL_META, SETTINGS, STYLE_META, WEB_SETTINGS};
    if is_bag(parent, name) {
        return schema::container_order(parent, name);
    }
    if let Some(o) = schema::nested(parent, name) {
        return Some(o);
    }
    Some(match name {
        "w:settings" => SETTINGS,
        "w:webSettings" => WEB_SETTINGS,
        "w:style" | "w:tblStylePr" => STYLE_ORDER,
        "w:lvl" => LEVEL_ORDER,
        "w:abstractNum" => ABSTRACT_ORDER,
        "w:font" => FONT,
        _ => {
            let _ = (STYLE_META, LEVEL_META, ABSTRACT_META);
            return None;
        }
    })
}

const STYLE_ORDER: &[&str] = &[
    "w:name", "w:aliases", "w:basedOn", "w:next", "w:link", "w:autoRedefine",
    "w:hidden", "w:uiPriority", "w:semiHidden", "w:unhideWhenUsed", "w:qFormat",
    "w:locked", "w:personal", "w:personalCompose", "w:personalReply", "w:rsid",
    "w:pPr", "w:rPr", "w:tblPr", "w:trPr", "w:tcPr", "w:tblStylePr",
];
const LEVEL_ORDER: &[&str] = &[
    "w:start", "w:numFmt", "w:lvlRestart", "w:pStyle", "w:isLgl", "w:suff", "w:lvlText",
    "w:lvlPicBulletId", "w:legacy", "w:lvlJc", "w:pPr", "w:rPr",
];
const ABSTRACT_ORDER: &[&str] = &["w:nsid", "w:multiLevelType", "w:tmpl", "w:name", "w:styleLink", "w:numStyleLink", "w:lvl"];

/// Move a row height written straight under w:tr into its trPr.
fn hoist_row_height(e: &Element) -> Element {
    let mut out = e.clone();
    let heights: Vec<Element> = out.elements().filter(|c| c.name == "w:trHeight").cloned().collect();
    out.children.retain(|n| !matches!(n, Node::Element(c) if c.name == "w:trHeight"));
    let at = out.children.iter().position(|n| matches!(n, Node::Element(c) if c.name == "w:trPr"));
    let at = match at {
        Some(i) => i,
        None => {
            let i = out.children.iter().position(|n| matches!(n, Node::Element(c) if c.name != "w:tblPrEx")).unwrap_or(out.children.len());
            out.children.insert(i, Node::Element(Element::new("w:trPr")));
            i
        }
    };
    if let Node::Element(trpr) = &mut out.children[at] {
        trpr.children.retain(|n| !matches!(n, Node::Element(c) if c.name == "w:trHeight"));
        for h in heights {
            trpr.children.push(Node::Element(h));
        }
    }
    out
}

fn first_difference(a: &Element, b: &Element, path: &str) -> Option<(String, String)> {
    if a.name != b.name {
        return Some((path.to_string(), format!("element {} vs {}", a.name, b.name)));
    }
    if a.attrs != b.attrs {
        let only_a: Vec<_> = a.attrs.iter().filter(|x| !b.attrs.contains(x)).collect();
        let only_b: Vec<_> = b.attrs.iter().filter(|x| !a.attrs.contains(x)).collect();
        return Some((path.to_string(), format!("attrs {only_a:?} vs {only_b:?}")));
    }
    let mut counts: HashMap<&str, usize> = HashMap::new();
    for (i, (x, y)) in a.children.iter().zip(b.children.iter()).enumerate() {
        match (x, y) {
            (Node::Element(x), Node::Element(y)) => {
                let k = counts.entry(x.name.as_str()).or_insert(0);
                *k += 1;
                let here = format!("{path}/{}[{}]", x.name, *k);
                if let Some(d) = first_difference(x, y, &here) {
                    return Some(d);
                }
            }
            (Node::Text(x), Node::Text(y)) if x == y => {}
            (x, y) => return Some((format!("{path}#{i}"), format!("{} vs {}", short(x), short(y)))),
        }
    }
    if a.children.len() != b.children.len() {
        let extra = if a.children.len() > b.children.len() { &a.children[b.children.len()] } else { &b.children[a.children.len()] };
        return Some((path.to_string(), format!("{} vs {} children; first extra {}", a.children.len(), b.children.len(), short(extra))));
    }
    None
}

fn short(n: &Node) -> String {
    match n {
        Node::Text(t) => format!("text {:?}", t.chars().take(40).collect::<String>()),
        Node::Element(e) => format!("<{}>", e.name),
    }
}

/// Prefixes that markup compatibility names in a part (mc:Choice Requires,
/// mc:Ignorable) but the part never declares. Canonical comparison cannot see
/// these, yet Word refuses such a file.
pub fn undeclared_prefixes(pkg: &Package) -> Vec<(String, String)> {
    let mut out = Vec::new();
    for (name, bytes) in &pkg.files {
        if !name.ends_with(".xml") {
            continue;
        }
        let text = crate::package::decode(bytes);
        let mut declared: Vec<&str> = Vec::new();
        for (i, _) in text.match_indices("xmlns:") {
            let rest = &text[i + 6..];
            if let Some(end) = rest.find('=') {
                declared.push(&rest[..end]);
            }
        }
        for key in ["Requires=\"", "Ignorable=\""] {
            for (i, _) in text.match_indices(key) {
                let rest = &text[i + key.len()..];
                let Some(end) = rest.find('"') else { continue };
                for p in rest[..end].split_whitespace() {
                    if !declared.contains(&p) {
                        out.push((name.clone(), p.to_string()));
                    }
                }
            }
        }
    }
    out.sort();
    out.dedup();
    out
}

/// Compare the content of two packages: the main document (with everything
/// its relationships reach) and the note parts.
pub fn compare(a: &Package, b: &Package) -> Vec<Diff> {
    compare_with(a, b, false)
}

/// Compare as [`compare`] does, reading where runs begin and end as
/// meaning nothing: neighbouring runs with the same properties are one run,
/// and a run with nothing in it is not there. Word itself joins such runs
/// when it saves, and so does editing; what the document says is the same.
pub fn compare_content(a: &Package, b: &Package) -> Vec<Diff> {
    compare_with(a, b, true)
}

fn compare_with(a: &Package, b: &Package, runs_as_text: bool) -> Vec<Diff> {
    let mut diffs = Vec::new();
    let (Ok(ma), Ok(mb)) = (a.main_part(), b.main_part()) else {
        return vec![Diff { part: "package".into(), path: String::new(), detail: "no main part".into() }];
    };
    let mut ca = Canon::new(a);
    let mut cb = Canon::new(b);
    ca.runs_as_text = runs_as_text;
    cb.runs_as_text = runs_as_text;
    let (Some(da), Some(db)) = (ca.part(&ma), cb.part(&mb)) else {
        return vec![Diff { part: ma, path: String::new(), detail: "unreadable".into() }];
    };
    if let Some((path, detail)) = first_difference(&da, &db, "") {
        // A differing linked part shows only as a different digest; name the
        // part and compare it so the report points inside it.
        diffs.push(Diff { part: "document".into(), path: path.clone(), detail: detail.clone() });
        if detail.contains("header:") || detail.contains("footer:") {
            diffs.extend(compare_linked(a, &ma, b, &mb, runs_as_text));
        }
    }
    // Notes, and every document-wide part the main document links to
    // (styles, numbering, settings, theme, font table, comments, ...).
    let mut kinds: Vec<String> = vec![REL_FOOTNOTES.to_string(), REL_ENDNOTES.to_string()];
    for r in a.rels(&ma).into_iter().chain(b.rels(&mb)) {
        if !r.external && !crate::package::is_content_rel(&r.rel_type) && !kinds.contains(&r.rel_type) {
            kinds.push(r.rel_type);
        }
    }
    for rel_type in &kinds {
        // A link to a part the package does not hold says nothing.
        let pa = a.rels(&ma).into_iter().find(|r| &r.rel_type == rel_type && a.files.contains_key(&r.target)).map(|r| r.target);
        let pb = b.rels(&mb).into_iter().find(|r| &r.rel_type == rel_type && b.files.contains_key(&r.target)).map(|r| r.target);
        let kind = rel_type.rsplit('/').next().unwrap().to_string();
        match (pa, pb) {
            (None, None) => {}
            (Some(x), Some(y)) => {
                let ex = ca.part(&x);
                let ey = cb.part(&y);
                match (ex, ey) {
                    (Some(ex), Some(ey)) => {
                        if let Some((path, detail)) = first_difference(&ex, &ey, "") {
                            diffs.push(Diff { part: kind, path, detail });
                        }
                    }
                    // Binary parts (a VBA project, fonts) compare as bytes.
                    _ => {
                        if a.files.get(&x) != b.files.get(&y) {
                            diffs.push(Diff { part: kind, path: String::new(), detail: "bytes differ".into() });
                        }
                    }
                }
            }
            (x, y) => diffs.push(Diff { part: kind, path: String::new(), detail: format!("present {} vs {}", x.is_some(), y.is_some()) }),
        }
    }
    diffs
}

/// Compare header and footer parts pairwise in reference order.
fn compare_linked(a: &Package, ma: &str, b: &Package, mb: &str, runs_as_text: bool) -> Vec<Diff> {
    let mut out = Vec::new();
    let refs = |p: &Package, m: &str| -> Vec<String> {
        let rels: HashMap<String, Rel> = p.rels(m).into_iter().map(|r| (r.id.clone(), r)).collect();
        let Ok(root) = p.xml(m) else { return vec![] };
        let mut all = Vec::new();
        root.walk(&mut all);
        all.into_iter()
            .filter(|e| e.name == "w:headerReference" || e.name == "w:footerReference")
            .filter_map(|e| e.attr("r:id").and_then(|id| rels.get(id)).map(|r| r.target.clone()))
            .collect()
    };
    let ra = refs(a, ma);
    let rb = refs(b, mb);
    let mut ca = Canon::new(a);
    let mut cb = Canon::new(b);
    ca.runs_as_text = runs_as_text;
    cb.runs_as_text = runs_as_text;
    for (x, y) in ra.iter().zip(rb.iter()) {
        if let (Some(ex), Some(ey)) = (ca.part(x), cb.part(y)) {
            if let Some((path, detail)) = first_difference(&ex, &ey, "") {
                out.push(Diff { part: x.clone(), path, detail });
                break;
            }
        }
    }
    out
}

/// Join neighbouring runs with the same properties, drop runs with nothing
/// in them, and join neighbouring text of one kind inside a run.
fn join_runs(children: Vec<Node>) -> Vec<Node> {
    let props = |e: &Element| e.child("w:rPr").cloned();
    let mut out: Vec<Node> = Vec::with_capacity(children.len());
    for n in children {
        let Node::Element(mut r) = n else {
            out.push(n);
            continue;
        };
        if r.name != "w:r" {
            out.push(Node::Element(r));
            continue;
        }
        if r.elements().all(|c| c.name == "w:rPr") {
            continue;
        }
        if let Some(Node::Element(prev)) = out.last_mut() {
            if prev.name == "w:r" && props(prev) == props(&r) {
                prev.children.extend(r.children.drain(..).filter(|c| !matches!(c, Node::Element(e) if e.name == "w:rPr")));
                join_text(prev);
                continue;
            }
        }
        join_text(&mut r);
        out.push(Node::Element(r));
    }
    out
}

fn join_text(r: &mut Element) {
    let mut out: Vec<Node> = Vec::with_capacity(r.children.len());
    for n in std::mem::take(&mut r.children) {
        if let (Some(Node::Element(prev)), Node::Element(e)) = (out.last_mut(), &n) {
            if prev.name == e.name && matches!(e.name.as_str(), "w:t" | "w:delText" | "w:instrText" | "w:delInstrText") {
                let text: String = e.children.iter().filter_map(|c| if let Node::Text(t) = c { Some(t.as_str()) } else { None }).collect();
                match prev.children.last_mut() {
                    Some(Node::Text(t)) => t.push_str(&text),
                    _ if text.is_empty() => {}
                    _ => prev.children.push(Node::Text(text)),
                }
                continue;
            }
        }
        out.push(n);
    }
    r.children = out;
}
