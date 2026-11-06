// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! Schema-checked property bags: pPr, rPr, tblPr, trPr, tcPr, tblPrEx, sectPr.
//!
//! A bag holds only children its container's schema names, recursively, and
//! always writes them in schema order, so a caller may set properties in any
//! order and still produce a file Word accepts.

use serde::{Deserialize, Serialize};

use crate::schema;
use crate::xml::{Element, Node};

/// A feature the model cannot express, named by its path.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Unsupported {
    pub path: String,
}

impl std::fmt::Display for Unsupported {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        write!(f, "unsupported: {}", self.path)
    }
}

impl std::error::Error for Unsupported {}

pub fn unsupported<T>(path: impl Into<String>) -> Result<T, Unsupported> {
    Err(Unsupported { path: path.into() })
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(transparent)]
pub struct Props(Element);

impl Props {
    /// An empty bag for a container (`"w:pPr"`, `"w:rPr"`, ...).
    pub fn new(name: &str) -> Self {
        Props(Element::new(name))
    }

    /// Check and order an element read from a file. `parent` is the element
    /// that holds it, because the same name means different bags in
    /// different places (a paragraph mark's rPr is not a run's).
    pub fn read(parent: &str, e: &Element) -> Result<Props, Unsupported> {
        let order = schema::container_order(parent, &e.name)
            .ok_or_else(|| Unsupported { path: format!("{parent}/{}", e.name) })?;
        Ok(Props(check(e, order, &e.name)?))
    }

    /// Validate a bag a caller built (for example deserialized from JSON).
    pub fn validated(self, parent: &str) -> Result<Props, Unsupported> {
        Props::read(parent, &self.0)
    }

    pub fn element(&self) -> &Element {
        &self.0
    }

    /// For rewriting attribute values in place (relationship ids). Adding or
    /// removing children here would bypass the schema order.
    pub(crate) fn element_mut(&mut self) -> &mut Element {
        &mut self.0
    }

    pub fn into_element(self) -> Element {
        self.0
    }

    pub fn name(&self) -> &str {
        &self.0.name
    }

    pub fn is_empty(&self) -> bool {
        self.0.children.is_empty() && self.0.attrs.is_empty()
    }

    pub fn get(&self, name: &str) -> Option<&Element> {
        self.0.child(name)
    }

    pub fn get_mut(&mut self, name: &str) -> Option<&mut Element> {
        self.0.elements_mut().find(|e| e.name == name)
    }

    /// Replace every child of this name with `child`, keeping schema order.
    pub fn set(&mut self, child: Element) -> Result<(), Unsupported> {
        let parent = self.0.name.clone();
        let order = self.order();
        let single = Element { name: parent.clone(), attrs: vec![], children: vec![Node::Element(child.clone())] };
        check(&single, order, &parent)?;
        self.remove(&child.name);
        self.0.children.push(Node::Element(child));
        sort(&mut self.0, order);
        Ok(())
    }

    /// Add a child alongside existing ones of the same name (tabs, columns,
    /// header references).
    pub fn push(&mut self, child: Element) -> Result<(), Unsupported> {
        let parent = self.0.name.clone();
        let order = self.order();
        let single = Element { name: parent.clone(), attrs: vec![], children: vec![Node::Element(child.clone())] };
        check(&single, order, &parent)?;
        self.0.children.push(Node::Element(child));
        sort(&mut self.0, order);
        Ok(())
    }

    pub fn remove(&mut self, name: &str) {
        self.0.children.retain(|n| !matches!(n, Node::Element(e) if e.name == name));
    }

    /// Children in schema order.
    pub fn children(&self) -> impl Iterator<Item = &Element> {
        self.0.elements()
    }

    pub fn retain(&mut self, f: impl FnMut(&Element) -> bool) {
        let mut f = f;
        self.0.children.retain(|n| match n {
            Node::Element(e) => f(e),
            Node::Text(_) => false,
        });
    }

    fn order(&self) -> &'static [&'static str] {
        match self.0.name.as_str() {
            "w:pPr" => schema::PPR,
            "w:rPr" => schema::RPR,
            "w:tblPr" => schema::TBLPR,
            "w:tblPrEx" => schema::TBLPREX,
            "w:trPr" => schema::TRPR,
            "w:tcPr" => schema::TCPR,
            "w:sectPr" => schema::SECTPR,
            _ => &[],
        }
    }
}

pub(crate) fn index(order: &[&str], name: &str) -> usize {
    // Header and footer references are one choice group (EG_HdrFtrReferences):
    // Word interleaves them, so they share a slot and keep their order.
    let name = if name == "w:footerReference" { "w:headerReference" } else { name };
    order.iter().position(|n| *n == name).unwrap_or(order.len())
}

fn sort(e: &mut Element, order: &[&str]) {
    e.children.sort_by_key(|n| match n {
        Node::Element(c) => index(order, &c.name),
        Node::Text(_) => usize::MAX,
    });
}

/// Attributes that only record editing sessions; never meaningful.
pub fn is_session_attr(name: &str) -> bool {
    name.starts_with("w:rsid") || name == "w14:paraId" || name == "w14:textId"
}

fn check(e: &Element, order: &'static [&'static str], path: &str) -> Result<Element, Unsupported> {
    let mut out = Element {
        name: e.name.clone(),
        attrs: e.attrs.iter().filter(|(k, _)| !is_session_attr(k)).cloned().collect(),
        children: Vec::new(),
    };
    for node in &e.children {
        let c = match node {
            Node::Element(c) => c,
            Node::Text(t) if t.trim().is_empty() => continue,
            Node::Text(_) => return unsupported(format!("{path}/#text")),
        };
        let here = format!("{path}/{}", c.name);
        if schema::is_extension(&c.name) {
            out.children.push(Node::Element(c.clone()));
            continue;
        }
        if !order.contains(&c.name.as_str()) {
            return unsupported(here);
        }
        let checked = if let Some(sub) = schema::container_order(&e.name, &c.name).filter(|_| is_bag(&e.name, &c.name)) {
            check(c, sub, &here)?
        } else if let Some(sub) = schema::nested(&e.name, &c.name) {
            check(c, sub, &here)?
        } else {
            if c.elements().next().is_some() {
                return unsupported(format!("{here}/{}", c.elements().next().unwrap().name));
            }
            Element {
                name: c.name.clone(),
                attrs: c.attrs.iter().filter(|(k, _)| !is_session_attr(k)).cloned().collect(),
                children: Vec::new(),
            }
        };
        out.children.push(Node::Element(checked));
    }
    sort(&mut out, order);
    Ok(out)
}

/// Whether `child` inside `parent` is itself a whole property bag (as a
/// paragraph mark's rPr inside pPr), rather than a nested group of leaves.
fn is_bag(parent: &str, child: &str) -> bool {
    matches!(
        (parent, child),
        ("w:pPr", "w:rPr") | ("w:pPr", "w:sectPr") | ("w:rPrChange", "w:rPr") | ("w:pPrChange", "w:pPr")
            | ("w:sectPrChange", "w:sectPr") | ("w:tblPrChange", "w:tblPr") | ("w:tblPrExChange", "w:tblPrEx")
            | ("w:trPrChange", "w:trPr") | ("w:tcPrChange", "w:tcPr")
    )
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::xml::parse;

    const W: &str = r#"xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main""#;

    #[test]
    fn writes_properties_in_schema_order_whatever_order_they_came() {
        let e = parse(&format!(r#"<w:pPr {W}><w:jc w:val="center"/><w:ind w:leftChars="100"/><w:pStyle w:val="A"/></w:pPr>"#)).unwrap();
        let p = Props::read("w:p", &e).unwrap();
        let names: Vec<_> = p.children().map(|c| c.name.as_str()).collect();
        assert_eq!(names, ["w:pStyle", "w:ind", "w:jc"]);
    }

    #[test]
    fn refuses_what_the_schema_does_not_name() {
        let e = parse(&format!(r#"<w:rPr {W}><w:b/><w:nonsense/></w:rPr>"#)).unwrap();
        assert_eq!(Props::read("w:r", &e).unwrap_err().path, "w:rPr/w:nonsense");
    }

    #[test]
    fn a_paragraph_mark_has_its_own_run_bag() {
        let e = parse(&format!(r#"<w:pPr {W}><w:rPr><w:sz w:val="21"/><w:ins w:id="1" w:author="a"/></w:rPr><w:tabs><w:tab w:val="left" w:pos="840"/></w:tabs></w:pPr>"#)).unwrap();
        let p = Props::read("w:p", &e).unwrap();
        let rpr = p.get("w:rPr").unwrap();
        let names: Vec<_> = rpr.elements().map(|c| c.name.as_str()).collect();
        assert_eq!(names, ["w:ins", "w:sz"]);
        assert_eq!(p.children().next().unwrap().name, "w:tabs");
    }

    #[test]
    fn set_replaces_and_keeps_order() {
        let mut p = Props::new("w:rPr");
        p.set(Element::new("w:sz").with_attr("w:val", "24")).unwrap();
        p.set(Element::new("w:b")).unwrap();
        p.set(Element::new("w:sz").with_attr("w:val", "28")).unwrap();
        let names: Vec<_> = p.children().map(|c| (c.name.as_str(), c.attr("w:val"))).collect();
        assert_eq!(names, [("w:b", None), ("w:sz", Some("28"))]);
        assert!(p.set(Element::new("w:jc")).is_err());
    }
}
