// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! A small owned XML tree with canonical namespace prefixes.
//!
//! Word files may bind WordprocessingML to any prefix, so names are rewritten
//! to one fixed prefix per namespace (`w:`, `r:`, `wp:` ...) on the way in.
//! Everything downstream compares and writes those canonical names.

use serde::{Deserialize, Serialize};

/// Namespace URI → the prefix this crate always uses for it.
pub const NAMESPACES: &[(&str, &str)] = &[
    ("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main"),
    ("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships"),
    ("wp", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"),
    ("a", "http://schemas.openxmlformats.org/drawingml/2006/main"),
    ("pic", "http://schemas.openxmlformats.org/drawingml/2006/picture"),
    ("c", "http://schemas.openxmlformats.org/drawingml/2006/chart"),
    ("dgm", "http://schemas.openxmlformats.org/drawingml/2006/diagram"),
    ("m", "http://schemas.openxmlformats.org/officeDocument/2006/math"),
    ("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006"),
    ("v", "urn:schemas-microsoft-com:vml"),
    ("o", "urn:schemas-microsoft-com:office:office"),
    ("w10", "urn:schemas-microsoft-com:office:word"),
    ("x", "urn:schemas-microsoft-com:office:excel"),
    ("wne", "http://schemas.microsoft.com/office/word/2006/wordml"),
    ("w14", "http://schemas.microsoft.com/office/word/2010/wordml"),
    ("w15", "http://schemas.microsoft.com/office/word/2012/wordml"),
    ("w16se", "http://schemas.microsoft.com/office/word/2015/wordml/symex"),
    ("w16cid", "http://schemas.microsoft.com/office/word/2016/wordml/cid"),
    ("w16", "http://schemas.microsoft.com/office/word/2018/wordml"),
    ("w16cex", "http://schemas.microsoft.com/office/word/2018/wordml/cex"),
    ("w16sdtdh", "http://schemas.microsoft.com/office/word/2020/wordml/sdtdatahash"),
    ("w16sdtfl", "http://schemas.microsoft.com/office/word/2024/wordml/sdtformatlock"),
    ("w16du", "http://schemas.microsoft.com/office/word/2023/wordml/word16du"),
    ("wp14", "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing"),
    ("wps", "http://schemas.microsoft.com/office/word/2010/wordprocessingShape"),
    ("wpg", "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"),
    ("wpc", "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas"),
    ("wpi", "http://schemas.microsoft.com/office/word/2010/wordprocessingInk"),
    ("a14", "http://schemas.microsoft.com/office/drawing/2010/main"),
    ("cx", "http://schemas.microsoft.com/office/drawing/2014/chartex"),
    ("aink", "http://schemas.microsoft.com/office/drawing/2016/ink"),
    ("am3d", "http://schemas.microsoft.com/office/drawing/2017/model3d"),
    ("sl", "http://schemas.openxmlformats.org/schemaLibrary/2006/main"),
    ("xml", "http://www.w3.org/XML/1998/namespace"),
    ("ve", "http://schemas.openxmlformats.org/markup-compatibility/2006"),
];

/// The canonical prefix for a namespace URI, if the crate knows it.
pub fn prefix_for(uri: &str) -> Option<&'static str> {
    NAMESPACES.iter().find(|(_, u)| *u == uri).map(|(p, _)| *p)
}

/// The URI for a canonical prefix.
pub fn uri_for(prefix: &str) -> Option<&'static str> {
    NAMESPACES.iter().find(|(p, _)| *p == prefix).map(|(_, u)| *u)
}

/// An XML element with canonical names. Attribute order is kept as read but
/// never compared.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct Element {
    pub name: String,
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub attrs: Vec<(String, String)>,
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub children: Vec<Node>,
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(untagged)]
pub enum Node {
    Element(Element),
    Text(String),
}

impl Element {
    pub fn new(name: impl Into<String>) -> Self {
        Element { name: name.into(), attrs: Vec::new(), children: Vec::new() }
    }

    pub fn with_attr(mut self, name: impl Into<String>, value: impl Into<String>) -> Self {
        self.set_attr(name, value);
        self
    }

    pub fn with_child(mut self, child: Element) -> Self {
        self.children.push(Node::Element(child));
        self
    }

    pub fn attr(&self, name: &str) -> Option<&str> {
        self.attrs.iter().find(|(k, _)| k == name).map(|(_, v)| v.as_str())
    }

    pub fn set_attr(&mut self, name: impl Into<String>, value: impl Into<String>) {
        let name = name.into();
        let value = value.into();
        if let Some(slot) = self.attrs.iter_mut().find(|(k, _)| *k == name) {
            slot.1 = value;
        } else {
            self.attrs.push((name, value));
        }
    }

    pub fn remove_attr(&mut self, name: &str) -> Option<String> {
        let at = self.attrs.iter().position(|(k, _)| k == name)?;
        Some(self.attrs.remove(at).1)
    }

    pub fn elements(&self) -> impl Iterator<Item = &Element> {
        self.children.iter().filter_map(|n| match n {
            Node::Element(e) => Some(e),
            Node::Text(_) => None,
        })
    }

    pub fn elements_mut(&mut self) -> impl Iterator<Item = &mut Element> {
        self.children.iter_mut().filter_map(|n| match n {
            Node::Element(e) => Some(e),
            Node::Text(_) => None,
        })
    }

    pub fn child(&self, name: &str) -> Option<&Element> {
        self.elements().find(|e| e.name == name)
    }

    /// Concatenated text of direct text children.
    pub fn text(&self) -> String {
        self.children
            .iter()
            .filter_map(|n| match n {
                Node::Text(t) => Some(t.as_str()),
                Node::Element(_) => None,
            })
            .collect()
    }

    /// Every element in document order, this one included.
    pub fn walk<'a>(&'a self, out: &mut Vec<&'a Element>) {
        out.push(self);
        for e in self.elements() {
            e.walk(out);
        }
    }

    /// The canonical prefixes used anywhere in this subtree, including those
    /// markup compatibility names in attribute VALUES (mc:Choice Requires,
    /// mc:Ignorable): a consumer resolves those too, and an undeclared one
    /// makes Word refuse the file.
    pub fn prefixes(&self, out: &mut std::collections::BTreeSet<String>) {
        if let Some((p, _)) = self.name.split_once(':') {
            out.insert(p.to_string());
        }
        for (k, v) in &self.attrs {
            if let Some((p, _)) = k.split_once(':') {
                if p != "xmlns" {
                    out.insert(p.to_string());
                }
            }
            let names_prefixes = matches!(
                k.as_str(),
                "Requires" | "mc:Ignorable" | "mc:ProcessContent" | "mc:PreserveElements" | "mc:PreserveAttributes" | "mc:MustUnderstand"
            );
            if names_prefixes {
                for word in v.split_whitespace() {
                    let p = word.split(':').next().unwrap_or(word);
                    if uri_for(p).is_some() {
                        out.insert(p.to_string());
                    }
                }
            }
        }
        for e in self.elements() {
            e.prefixes(out);
        }
    }
}

/// Elements whose text content is meaningful character data. Whitespace-only
/// text anywhere else is formatting and is dropped.
fn keeps_text(name: &str) -> bool {
    matches!(
        name,
        "w:t" | "w:instrText" | "w:delText" | "w:delInstrText" | "m:t" | "a:t"
            | "wp:posOffset" | "wp:align" | "w:tblCaption" | "w:tblDescription"
    ) || !name.starts_with("w:")
}

#[derive(Debug)]
pub struct XmlError(pub String);

impl std::fmt::Display for XmlError {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        f.write_str(&self.0)
    }
}

impl std::error::Error for XmlError {}

/// Parse XML text into an owned tree with canonical prefixes.
pub fn parse(text: &str) -> Result<Element, XmlError> {
    let options = roxmltree::ParsingOptions { allow_dtd: true, ..Default::default() };
    let doc = roxmltree::Document::parse_with_options(text, options)
        .map_err(|e| XmlError(e.to_string()))?;
    Ok(convert(doc.root_element()))
}

fn qualified(ns: Option<&str>, local: &str, node: roxmltree::Node) -> String {
    match ns {
        None => local.to_string(),
        Some(uri) => match prefix_for(uri) {
            Some(p) => format!("{p}:{local}"),
            None => {
                // An unknown namespace keeps whatever prefix the file gave it.
                let p = node
                    .namespaces()
                    .find(|n| n.uri() == uri)
                    .and_then(|n| n.name())
                    .unwrap_or("ns");
                format!("{{{uri}}}{p}:{local}")
            }
        },
    }
}

fn convert(node: roxmltree::Node) -> Element {
    let tag = node.tag_name();
    let name = qualified(tag.namespace(), tag.name(), node);
    let attrs = node
        .attributes()
        .map(|a| (qualified(a.namespace(), a.name(), node), a.value().to_string()))
        .collect();
    let keep = keeps_text(&name);
    let mut children = Vec::new();
    for ch in node.children() {
        if ch.is_element() {
            children.push(Node::Element(convert(ch)));
        } else if ch.is_text() {
            let t = ch.text().unwrap_or("");
            if keep || !t.trim().is_empty() {
                children.push(Node::Text(t.to_string()));
            }
        }
    }
    Element { name, attrs, children }
}

pub fn escape(s: &str, attr: bool) -> String {
    let mut out = String::with_capacity(s.len());
    for c in s.chars() {
        match c {
            '&' => out.push_str("&amp;"),
            '<' => out.push_str("&lt;"),
            '>' => out.push_str("&gt;"),
            '"' if attr => out.push_str("&quot;"),
            '\t' if attr => out.push_str("&#9;"),
            '\n' if attr => out.push_str("&#10;"),
            '\r' => out.push_str("&#13;"),
            c => out.push(c),
        }
    }
    out
}

/// Strip the `{uri}` marker an unknown-namespace name carries.
fn plain(name: &str) -> &str {
    match name.find('}') {
        Some(i) if name.starts_with('{') => &name[i + 1..],
        _ => name,
    }
}

fn write_into(e: &Element, out: &mut String) {
    out.push('<');
    out.push_str(plain(&e.name));
    for (k, v) in &e.attrs {
        out.push(' ');
        out.push_str(plain(k));
        out.push_str("=\"");
        out.push_str(&escape(v, true));
        out.push('"');
    }
    if e.children.is_empty() {
        out.push_str("/>");
        return;
    }
    out.push('>');
    for ch in &e.children {
        match ch {
            Node::Element(c) => write_into(c, out),
            Node::Text(t) => out.push_str(&escape(t, false)),
        }
    }
    out.push_str("</");
    out.push_str(plain(&e.name));
    out.push('>');
}

/// Serialize an element. `root` adds the XML declaration and declares every
/// namespace the subtree uses.
pub fn write(e: &Element, root: bool) -> String {
    let mut out = String::new();
    if root {
        out.push_str("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n");
        let mut el = e.clone();
        let mut used = std::collections::BTreeSet::new();
        e.prefixes(&mut used);
        let mut unknown = Vec::new();
        e.walk_unknown(&mut unknown);
        for p in &used {
            if p == "xml" {
                continue;
            }
            if let Some(uri) = uri_for(p) {
                if el.attr(&format!("xmlns:{p}")).is_none() {
                    el.attrs.push((format!("xmlns:{p}"), uri.to_string()));
                }
            }
        }
        for (p, uri) in unknown {
            if el.attr(&format!("xmlns:{p}")).is_none() {
                el.attrs.push((format!("xmlns:{p}"), uri));
            }
        }
        write_into(&el, &mut out);
    } else {
        write_into(e, &mut out);
    }
    out
}

impl Element {
    fn walk_unknown(&self, out: &mut Vec<(String, String)>) {
        let mut note = |name: &str| {
            if let Some(rest) = name.strip_prefix('{') {
                if let Some((uri, local)) = rest.split_once('}') {
                    if let Some((p, _)) = local.split_once(':') {
                        if !out.iter().any(|(q, _)| q == p) {
                            out.push((p.to_string(), uri.to_string()));
                        }
                    }
                }
            }
        };
        note(&self.name);
        for (k, _) in &self.attrs {
            note(k);
        }
        for e in self.elements() {
            e.walk_unknown(out);
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn rewrites_any_prefix_to_the_canonical_one() {
        let e = parse(r#"<x:document xmlns:x="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><x:body><x:p/></x:body></x:document>"#).unwrap();
        assert_eq!(e.name, "w:document");
        assert_eq!(e.elements().next().unwrap().name, "w:body");
        let out = write(&e, true);
        assert!(out.contains(r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p/></w:body></w:document>"#), "{out}");
    }

    #[test]
    fn declares_a_prefix_only_named_in_a_requires_value() {
        // Word refuses a part whose mc:Choice requires an undeclared prefix.
        let e = parse(r#"<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"><mc:AlternateContent><mc:Choice Requires="wpg"><w:t/></mc:Choice></mc:AlternateContent></w:r>"#).unwrap();
        let out = write(&e, true);
        assert!(out.contains(r#"xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup""#), "{out}");
    }

    #[test]
    fn keeps_spaces_only_where_they_are_text() {
        let e = parse(r#"<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:r> <w:t xml:space="preserve">  a </w:t></w:r></w:p>"#).unwrap();
        let r = e.elements().next().unwrap();
        assert_eq!(r.children.len(), 1);
        assert_eq!(r.elements().next().unwrap().text(), "  a ");
        assert_eq!(r.elements().next().unwrap().attr("xml:space"), Some("preserve"));
    }
}
