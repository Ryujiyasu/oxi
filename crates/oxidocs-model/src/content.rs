// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! Document content: blocks, paragraphs, runs, tables and the objects they
//! carry. Every type here is plain data (serde), so a caller can build a
//! document as JSON and hand it to the writer.

use serde::{Deserialize, Serialize};

use crate::props::{is_session_attr, unsupported, Props, Unsupported};
use crate::xml::{Element, Node};

/// A part an object points at: a picture, a chart, an embedded workbook, or
/// an external link. Internal parts carry their own relationships, whose ids
/// their XML already uses.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub enum Part {
    Internal {
        rel_type: String,
        /// Suggested file name, for example `media/image1.png`.
        name: String,
        content_type: String,
        data: Vec<u8>,
        #[serde(default, skip_serializing_if = "Vec::is_empty")]
        rels: Vec<(String, Part)>,
    },
    External {
        rel_type: String,
        target: String,
    },
}

/// An object kept as its own markup: a DrawingML picture or shape, a VML
/// shape, an OLE object, an equation. Relationship attributes inside `xml`
/// are written as `#N`, the index of the part in `parts`.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct Object {
    pub xml: Element,
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub parts: Vec<Part>,
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub enum Block {
    Paragraph(Paragraph),
    Table(Table),
    Sdt(Sdt<Block>),
    CustomXml(Wrapper<Block>),
    Revision(Revision<Block>),
    Marker(Marker),
    /// An alternate-format chunk (altChunk) imported at this point.
    AltChunk(Object),
}

#[derive(Debug, Clone, PartialEq, Default, Serialize, Deserialize)]
pub struct Paragraph {
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub props: Option<Props>,
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub content: Vec<Inline>,
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub attrs: Vec<(String, String)>,
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub enum Inline {
    Run(Run),
    Hyperlink(Hyperlink),
    SimpleField(SimpleField),
    Sdt(Sdt<Inline>),
    SmartTag(Wrapper<Inline>),
    CustomXml(Wrapper<Inline>),
    Revision(Revision<Inline>),
    /// Bidirectional embedding (w:dir) and override (w:bdo).
    Bidi(Wrapper<Inline>),
    Marker(Marker),
    /// Office Math (m:oMath / m:oMathPara).
    Math(Object),
    /// Run content written straight into the paragraph, outside any run (a
    /// w:br some generators put between runs). Kept where it was written.
    Loose(RunItem),
}

#[derive(Debug, Clone, PartialEq, Default, Serialize, Deserialize)]
pub struct Run {
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub props: Option<Props>,
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub content: Vec<RunItem>,
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub enum RunItem {
    Text(String),
    DeletedText(String),
    InstrText(String),
    DeletedInstrText(String),
    Tab,
    /// w:br; `kind` is page, column or textWrapping, `clear` for wrapping.
    Break {
        #[serde(default, skip_serializing_if = "Vec::is_empty")]
        attrs: Vec<(String, String)>,
    },
    /// w:fldChar; `data` holds form-field data (w:ffData) when present.
    FieldChar {
        attrs: Vec<(String, String)>,
        #[serde(default, skip_serializing_if = "Option::is_none")]
        data: Option<Element>,
    },
    /// Leaf run content named by `name`: cr, noBreakHyphen, softHyphen, sym,
    /// ptab, pgNum, footnoteReference, endnoteRef, separator, ...
    Mark {
        name: String,
        #[serde(default, skip_serializing_if = "Vec::is_empty")]
        attrs: Vec<(String, String)>,
    },
    Ruby(Ruby),
    /// w:drawing, w:pict, w:object, mc:AlternateContent, w:contentPart.
    Object(Object),
    /// Run properties written out of place: after the content, or a second
    /// rPr in the same run.
    Props(Props),
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct Ruby {
    pub props: Element,
    pub text: Vec<Inline>,
    pub base: Vec<Inline>,
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct Hyperlink {
    /// An external target (the relationship's URL), if any.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub target: Option<String>,
    /// w:anchor, w:tooltip, w:history, w:tgtFrame, w:docLocation.
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub attrs: Vec<(String, String)>,
    pub content: Vec<Inline>,
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct SimpleField {
    pub attrs: Vec<(String, String)>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub data: Option<Element>,
    pub content: Vec<Inline>,
}

/// A content control: its settings (w:sdtPr), end-mark settings and content.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct Sdt<T> {
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub props: Option<Element>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub end_props: Option<Element>,
    pub content: Vec<T>,
}

/// A container that adds attributes or a property element to its content:
/// customXml, smartTag, dir/bdo.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct Wrapper<T> {
    pub name: String,
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub attrs: Vec<(String, String)>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub props: Option<Element>,
    pub content: Vec<T>,
}

/// A tracked insertion, deletion or move around content.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct Revision<T> {
    /// w:ins, w:del, w:moveFrom, w:moveTo.
    pub name: String,
    pub attrs: Vec<(String, String)>,
    pub content: Vec<T>,
}

/// A point marker: bookmarks, comment ranges, permission ranges, move ranges.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct Marker {
    pub name: String,
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub attrs: Vec<(String, String)>,
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct Table {
    pub props: Props,
    /// Grid column widths as written (w:gridCol w:w), with any grid change.
    /// Some generators leave the grid out; then so does the model.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub grid: Option<Element>,
    pub rows: Vec<RowItem>,
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub enum RowItem {
    Row(Row),
    Sdt(Sdt<RowItem>),
    CustomXml(Wrapper<RowItem>),
    Revision(Revision<RowItem>),
    Marker(Marker),
}

#[derive(Debug, Clone, PartialEq, Default, Serialize, Deserialize)]
pub struct Row {
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub props_ex: Option<Props>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub props: Option<Props>,
    pub cells: Vec<CellItem>,
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub attrs: Vec<(String, String)>,
    /// The file wrote the row properties after the cells (some generators
    /// do); kept that way.
    #[serde(default, skip_serializing_if = "std::ops::Not::not")]
    pub props_after_cells: bool,
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub enum CellItem {
    Cell(Cell),
    Sdt(Sdt<CellItem>),
    CustomXml(Wrapper<CellItem>),
    Revision(Revision<CellItem>),
    Marker(Marker),
}

#[derive(Debug, Clone, PartialEq, Default, Serialize, Deserialize)]
pub struct Cell {
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub props: Option<Props>,
    pub content: Vec<Block>,
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub attrs: Vec<(String, String)>,
}

// ---------------------------------------------------------------------------
// Reading
// ---------------------------------------------------------------------------

/// What content reading needs from the package: turning a relationship id
/// into the part it names.
pub trait Resolver {
    fn part(&mut self, rid: &str) -> Result<Part, Unsupported>;
}

const MARKERS: &[&str] = &[
    "w:bookmarkStart", "w:bookmarkEnd", "w:commentRangeStart", "w:commentRangeEnd",
    "w:permStart", "w:permEnd", "w:moveFromRangeStart", "w:moveFromRangeEnd",
    "w:moveToRangeStart", "w:moveToRangeEnd", "w:customXmlInsRangeStart",
    "w:customXmlInsRangeEnd", "w:customXmlDelRangeStart", "w:customXmlDelRangeEnd",
];

/// Things that record only an editing session or a past layout.
const DROPPED: &[&str] = &["w:proofErr", "w:lastRenderedPageBreak"];

const RUN_MARKS: &[&str] = &[
    "w:cr", "w:noBreakHyphen", "w:softHyphen", "w:sym", "w:ptab", "w:pgNum",
    "w:footnoteReference", "w:endnoteReference", "w:commentReference",
    "w:footnoteRef", "w:endnoteRef", "w:annotationRef", "w:separator",
    "w:continuationSeparator", "w:dayShort", "w:monthShort", "w:yearShort",
    "w:dayLong", "w:monthLong", "w:yearLong",
];

const SDT_PR: &[&str] = &[
    "w:rPr", "w:alias", "w:tag", "w:id", "w:lock", "w:placeholder", "w:temporary",
    "w:showingPlcHdr", "w:dataBinding", "w:label", "w:tabIndex", "w:equation",
    "w:comboBox", "w:date", "w:docPartObj", "w:docPartList", "w:dropDownList",
    "w:picture", "w:richText", "w:text", "w:citation", "w:group", "w:bibliography",
];

fn attrs(e: &Element) -> Vec<(String, String)> {
    e.attrs.iter().filter(|(k, _)| !is_session_attr(k)).cloned().collect()
}

fn text_of(e: &Element) -> String {
    e.text()
}

/// Children of a content-control settings element, checked by name; their
/// own subtrees are control configuration and are kept whole.
fn read_sdt_pr(e: &Element, path: &str) -> Result<Element, Unsupported> {
    for c in e.elements() {
        let ok = SDT_PR.contains(&c.name.as_str()) || crate::schema::is_extension(&c.name);
        if !ok {
            return unsupported(format!("{path}/{}", c.name));
        }
    }
    let mut out = e.clone();
    if let Some(rpr) = e.child("w:rPr") {
        let checked = Props::read("w:sdtPr", rpr).map(|p| p.into_element());
        let checked = checked.or_else(|_| Props::read("w:r", rpr).map(|p| p.into_element()))?;
        for n in out.children.iter_mut() {
            if let Node::Element(c) = n {
                if c.name == "w:rPr" {
                    *c = checked.clone();
                }
            }
        }
    }
    Ok(out)
}

/// Reads one child element into a list of items of the caller's kind.
type ReadItem<'a, T> = dyn FnMut(&Element, &str, &mut dyn Resolver, &mut Vec<T>) -> Result<(), Unsupported> + 'a;

fn read_sdt<T>(
    e: &Element,
    path: &str,
    res: &mut dyn Resolver,
    item: &mut ReadItem<T>,
) -> Result<Sdt<T>, Unsupported> {
    let mut sdt = Sdt { props: None, end_props: None, content: Vec::new() };
    for c in e.elements() {
        let here = format!("{path}/{}", c.name);
        match c.name.as_str() {
            "w:sdtPr" => sdt.props = Some(read_sdt_pr(c, &here)?),
            "w:sdtEndPr" => {
                let mut end = c.clone();
                if let Some(rpr) = c.child("w:rPr") {
                    let checked = Props::read("w:r", rpr)?.into_element();
                    end.children = vec![Node::Element(checked)];
                }
                sdt.end_props = Some(end);
            }
            "w:sdtContent" => {
                for x in c.elements() {
                    item(x, &here, res, &mut sdt.content)?;
                }
            }
            _ => return unsupported(here),
        }
    }
    Ok(sdt)
}

fn read_wrapper<T>(
    e: &Element,
    path: &str,
    props_name: &str,
    res: &mut dyn Resolver,
    item: &mut ReadItem<T>,
) -> Result<Wrapper<T>, Unsupported> {
    let mut w = Wrapper { name: e.name.clone(), attrs: attrs(e), props: None, content: Vec::new() };
    for c in e.elements() {
        if c.name == props_name {
            w.props = Some(c.clone());
        } else {
            item(c, path, res, &mut w.content)?;
        }
    }
    Ok(w)
}

fn read_revision<T>(
    e: &Element,
    path: &str,
    res: &mut dyn Resolver,
    item: &mut ReadItem<T>,
) -> Result<Revision<T>, Unsupported> {
    let mut r = Revision { name: e.name.clone(), attrs: attrs(e), content: Vec::new() };
    for c in e.elements() {
        item(c, path, res, &mut r.content)?;
    }
    Ok(r)
}

/// Collect every relationship an object's markup names and number them.
pub fn read_object(e: &Element, res: &mut dyn Resolver) -> Result<Object, Unsupported> {
    let mut parts: Vec<Part> = Vec::new();
    let mut seen: Vec<String> = Vec::new();
    let mut xml = e.clone();
    remap_rels(&mut xml, &mut |rid| {
        if let Some(i) = seen.iter().position(|s| s == rid) {
            return Ok(format!("#{i}"));
        }
        let part = match res.part(rid) {
            Ok(p) => p,
            // A reference to a relationship the part does not have: it
            // points nowhere, and so must the copy.
            Err(u) if u.path.ends_with("(dangling)") => return Ok(dangling(rid)),
            Err(u) => return Err(u),
        };
        seen.push(rid.to_string());
        parts.push(part);
        Ok(format!("#{}", parts.len() - 1))
    })?;
    strip_session(&mut xml);
    Ok(Object { xml, parts })
}

fn strip_session(e: &mut Element) {
    e.attrs.retain(|(k, _)| !is_session_attr(k));
    for c in e.elements_mut() {
        strip_session(c);
    }
}

/// The id written for a reference whose relationship is missing: one no
/// writer allocates, so it stays unresolved.
pub fn dangling(rid: &str) -> String {
    format!("oxiMissing_{rid}")
}

/// Attributes in the relationships namespace hold relationship ids.
fn is_rel_attr(name: &str) -> bool {
    name.starts_with("r:")
}

pub fn remap_rels(
    e: &mut Element,
    f: &mut dyn FnMut(&str) -> Result<String, Unsupported>,
) -> Result<(), Unsupported> {
    for (k, v) in e.attrs.iter_mut() {
        if is_rel_attr(k) && !v.is_empty() {
            *v = f(v)?;
        }
    }
    for c in e.elements_mut() {
        remap_rels(c, f)?;
    }
    Ok(())
}

pub fn read_blocks(parent: &Element, path: &str, res: &mut dyn Resolver) -> Result<Vec<Block>, Unsupported> {
    let mut out = Vec::new();
    for c in parent.elements() {
        read_block(c, path, res, &mut out)?;
    }
    Ok(out)
}

fn read_block(e: &Element, path: &str, res: &mut dyn Resolver, out: &mut Vec<Block>) -> Result<(), Unsupported> {
    let here = format!("{path}/{}", e.name);
    let name = e.name.as_str();
    if DROPPED.contains(&name) {
        return Ok(());
    }
    if MARKERS.contains(&name) {
        out.push(Block::Marker(Marker { name: e.name.clone(), attrs: attrs(e) }));
        return Ok(());
    }
    let block = match name {
        "w:p" => Block::Paragraph(read_paragraph(e, &here, res)?),
        "w:tbl" => Block::Table(read_table(e, &here, res)?),
        "w:sdt" => Block::Sdt(read_sdt(e, &here, res, &mut |c, p, r, v| read_block(c, p, r, v))?),
        "w:customXml" => Block::CustomXml(read_wrapper(e, &here, "w:customXmlPr", res, &mut |c, p, r, v| read_block(c, p, r, v))?),
        "w:ins" | "w:del" | "w:moveFrom" | "w:moveTo" => {
            Block::Revision(read_revision(e, &here, res, &mut |c, p, r, v| read_block(c, p, r, v))?)
        }
        "w:altChunk" => Block::AltChunk(read_object(e, res)?),
        _ => return unsupported(here),
    };
    out.push(block);
    Ok(())
}

pub fn read_paragraph(e: &Element, path: &str, res: &mut dyn Resolver) -> Result<Paragraph, Unsupported> {
    let mut p = Paragraph { props: None, content: Vec::new(), attrs: attrs(e) };
    for c in e.elements() {
        if c.name == "w:pPr" {
            p.props = Some(Props::read("w:p", c)?);
        } else {
            read_inline(c, path, res, &mut p.content)?;
        }
    }
    Ok(p)
}

fn read_inline(e: &Element, path: &str, res: &mut dyn Resolver, out: &mut Vec<Inline>) -> Result<(), Unsupported> {
    let here = format!("{path}/{}", e.name);
    let name = e.name.as_str();
    if DROPPED.contains(&name) {
        return Ok(());
    }
    if MARKERS.contains(&name) {
        out.push(Inline::Marker(Marker { name: e.name.clone(), attrs: attrs(e) }));
        return Ok(());
    }
    let item = match name {
        "w:r" => Inline::Run(read_run(e, &here, res)?),
        "w:hyperlink" => {
            let mut a = attrs(e);
            let mut target = None;
            if let Some(pos) = a.iter().position(|(k, _)| k == "r:id") {
                let rid = a.remove(pos).1;
                match res.part(&rid)? {
                    Part::External { target: t, .. } => target = Some(t),
                    Part::Internal { .. } => return unsupported(format!("{here}@r:id(internal)")),
                }
            }
            let mut content = Vec::new();
            for c in e.elements() {
                read_inline(c, &here, res, &mut content)?;
            }
            Inline::Hyperlink(Hyperlink { target, attrs: a, content })
        }
        "w:fldSimple" => {
            let mut f = SimpleField { attrs: attrs(e), data: None, content: Vec::new() };
            for c in e.elements() {
                if c.name == "w:fldData" {
                    f.data = Some(c.clone());
                } else {
                    read_inline(c, &here, res, &mut f.content)?;
                }
            }
            Inline::SimpleField(f)
        }
        "w:sdt" => Inline::Sdt(read_sdt(e, &here, res, &mut |c, p, r, v| read_inline(c, p, r, v))?),
        "w:smartTag" => Inline::SmartTag(read_wrapper(e, &here, "w:smartTagPr", res, &mut |c, p, r, v| read_inline(c, p, r, v))?),
        "w:customXml" => Inline::CustomXml(read_wrapper(e, &here, "w:customXmlPr", res, &mut |c, p, r, v| read_inline(c, p, r, v))?),
        "w:dir" | "w:bdo" => Inline::Bidi(read_wrapper(e, &here, "", res, &mut |c, p, r, v| read_inline(c, p, r, v))?),
        "w:ins" | "w:del" | "w:moveFrom" | "w:moveTo" => {
            Inline::Revision(read_revision(e, &here, res, &mut |c, p, r, v| read_inline(c, p, r, v))?)
        }
        "m:oMath" | "m:oMathPara" => Inline::Math(read_object(e, res)?),
        "w:rPr" => Inline::Loose(RunItem::Props(Props::read("w:r", e)?)),
        "w:br" | "w:tab" | "w:cr" | "w:t" => {
            let holder = Element { name: "w:r".into(), attrs: vec![], children: vec![Node::Element(e.clone())] };
            let mut run = read_run(&holder, path, res)?;
            Inline::Loose(run.content.remove(0))
        }
        _ => return unsupported(here),
    };
    out.push(item);
    Ok(())
}

pub fn read_run(e: &Element, path: &str, res: &mut dyn Resolver) -> Result<Run, Unsupported> {
    let mut run = Run::default();
    for c in e.elements() {
        let here = format!("{path}/{}", c.name);
        let name = c.name.as_str();
        if DROPPED.contains(&name) {
            continue;
        }
        let item = match name {
            "w:rPr" if run.props.is_none() && run.content.is_empty() => {
                run.props = Some(Props::read("w:r", c)?);
                continue;
            }
            // A second rPr, or one after the content: some generators write
            // these. Kept where the file put it.
            "w:rPr" => RunItem::Props(Props::read("w:r", c)?),
            "w:t" => RunItem::Text(text_of(c)),
            "w:delText" => RunItem::DeletedText(text_of(c)),
            "w:instrText" => RunItem::InstrText(text_of(c)),
            "w:delInstrText" => RunItem::DeletedInstrText(text_of(c)),
            "w:tab" => RunItem::Tab,
            "w:br" => RunItem::Break { attrs: attrs(c) },
            "w:fldChar" => {
                let data = match c.elements().next() {
                    None => None,
                    Some(d) if matches!(d.name.as_str(), "w:ffData" | "w:fldData" | "w:numberingChange") => Some(d.clone()),
                    Some(d) => return unsupported(format!("{here}/{}", d.name)),
                };
                RunItem::FieldChar { attrs: attrs(c), data }
            }
            "w:ruby" => RunItem::Ruby(read_ruby(c, &here, res)?),
            "w:drawing" | "w:pict" | "w:object" | "mc:AlternateContent" | "w:contentPart" => {
                RunItem::Object(read_object(c, res)?)
            }
            n if RUN_MARKS.contains(&n) => {
                if c.elements().next().is_some() {
                    return unsupported(format!("{here}/{}", c.elements().next().unwrap().name));
                }
                RunItem::Mark { name: c.name.clone(), attrs: attrs(c) }
            }
            _ => return unsupported(here),
        };
        run.content.push(item);
    }
    Ok(run)
}

const RUBY_PR: &[&str] = &["w:rubyAlign", "w:hps", "w:hpsRaise", "w:hpsBaseText", "w:lid", "w:dirty"];

fn read_ruby(e: &Element, path: &str, res: &mut dyn Resolver) -> Result<Ruby, Unsupported> {
    let mut props = None;
    let mut text = Vec::new();
    let mut base = Vec::new();
    for c in e.elements() {
        let here = format!("{path}/{}", c.name);
        match c.name.as_str() {
            "w:rubyPr" => {
                for x in c.elements() {
                    if !RUBY_PR.contains(&x.name.as_str()) {
                        return unsupported(format!("{here}/{}", x.name));
                    }
                }
                props = Some(c.clone());
            }
            "w:rt" => {
                for x in c.elements() {
                    read_inline(x, &here, res, &mut text)?;
                }
            }
            "w:rubyBase" => {
                for x in c.elements() {
                    read_inline(x, &here, res, &mut base)?;
                }
            }
            _ => return unsupported(here),
        }
    }
    Ok(Ruby { props: props.unwrap_or_else(|| Element::new("w:rubyPr")), text, base })
}

fn read_table(e: &Element, path: &str, res: &mut dyn Resolver) -> Result<Table, Unsupported> {
    let mut props = Props::new("w:tblPr");
    let mut grid = None;
    let mut rows = Vec::new();
    for c in e.elements() {
        match c.name.as_str() {
            "w:tblPr" => props = Props::read("w:tbl", c)?,
            "w:tblGrid" => {
                for g in c.elements() {
                    if g.name != "w:gridCol" && g.name != "w:tblGridChange" {
                        return unsupported(format!("{path}/w:tblGrid/{}", g.name));
                    }
                }
                let mut g = c.clone();
                strip_session(&mut g);
                grid = Some(g);
            }
            _ => read_row_item(c, path, res, &mut rows)?,
        }
    }
    Ok(Table { props, grid, rows })
}

fn read_row_item(e: &Element, path: &str, res: &mut dyn Resolver, out: &mut Vec<RowItem>) -> Result<(), Unsupported> {
    let here = format!("{path}/{}", e.name);
    let name = e.name.as_str();
    if DROPPED.contains(&name) {
        return Ok(());
    }
    if MARKERS.contains(&name) {
        out.push(RowItem::Marker(Marker { name: e.name.clone(), attrs: attrs(e) }));
        return Ok(());
    }
    let item = match name {
        "w:tr" => {
            let mut row = Row { attrs: attrs(e), ..Default::default() };
            for c in e.elements() {
                match c.name.as_str() {
                    "w:tblPrEx" => row.props_ex = Some(Props::read("w:tr", c)?),
                    "w:trPr" => {
                        row.props_after_cells = !row.cells.is_empty();
                        let mut props = Props::read("w:tr", c)?;
                        if let Some(old) = row.props.take() {
                            for x in old.children() {
                                props.push(x.clone())?;
                            }
                        }
                        row.props = Some(props);
                    }
                    // Generated files sometimes put the row height straight
                    // under w:tr; it is read as the row's height (as the
                    // layout engine reads it), so it joins trPr.
                    "w:trHeight" => {
                        let mut props = row.props.take().unwrap_or_else(|| Props::new("w:trPr"));
                        let mut h = c.clone();
                        h.attrs.retain(|(k, _)| !is_session_attr(k));
                        props.set(h)?;
                        row.props = Some(props);
                    }
                    _ => read_cell_item(c, &here, res, &mut row.cells)?,
                }
            }
            RowItem::Row(row)
        }
        "w:sdt" => RowItem::Sdt(read_sdt(e, &here, res, &mut |c, p, r, v| read_row_item(c, p, r, v))?),
        "w:customXml" => RowItem::CustomXml(read_wrapper(e, &here, "w:customXmlPr", res, &mut |c, p, r, v| read_row_item(c, p, r, v))?),
        "w:ins" | "w:del" | "w:moveFrom" | "w:moveTo" => {
            RowItem::Revision(read_revision(e, &here, res, &mut |c, p, r, v| read_row_item(c, p, r, v))?)
        }
        _ => return unsupported(here),
    };
    out.push(item);
    Ok(())
}

fn read_cell_item(e: &Element, path: &str, res: &mut dyn Resolver, out: &mut Vec<CellItem>) -> Result<(), Unsupported> {
    let here = format!("{path}/{}", e.name);
    let name = e.name.as_str();
    if DROPPED.contains(&name) {
        return Ok(());
    }
    if MARKERS.contains(&name) {
        out.push(CellItem::Marker(Marker { name: e.name.clone(), attrs: attrs(e) }));
        return Ok(());
    }
    let item = match name {
        "w:tc" => {
            let mut cell = Cell { attrs: attrs(e), ..Default::default() };
            for c in e.elements() {
                if c.name == "w:tcPr" {
                    cell.props = Some(Props::read("w:tc", c)?);
                } else {
                    read_block(c, &here, res, &mut cell.content)?;
                }
            }
            CellItem::Cell(cell)
        }
        "w:sdt" => CellItem::Sdt(read_sdt(e, &here, res, &mut |c, p, r, v| read_cell_item(c, p, r, v))?),
        "w:customXml" => CellItem::CustomXml(read_wrapper(e, &here, "w:customXmlPr", res, &mut |c, p, r, v| read_cell_item(c, p, r, v))?),
        "w:ins" | "w:del" | "w:moveFrom" | "w:moveTo" => {
            CellItem::Revision(read_revision(e, &here, res, &mut |c, p, r, v| read_cell_item(c, p, r, v))?)
        }
        _ => return unsupported(here),
    };
    out.push(item);
    Ok(())
}

// ---------------------------------------------------------------------------
// Writing
// ---------------------------------------------------------------------------

/// What content writing needs from the package: a relationship id for a part.
pub trait PartSink {
    fn add(&mut self, part: &Part) -> String;
}

fn el(name: &str, attrs: &[(String, String)]) -> Element {
    Element { name: name.to_string(), attrs: attrs.to_vec(), children: Vec::new() }
}

pub fn write_object(o: &Object, sink: &mut dyn PartSink) -> Element {
    let rids: Vec<String> = o.parts.iter().map(|p| sink.add(p)).collect();
    let mut xml = o.xml.clone();
    let _ = remap_rels(&mut xml, &mut |v| {
        let i: Option<usize> = v.strip_prefix('#').and_then(|n| n.parse().ok());
        Ok(i.and_then(|i| rids.get(i).cloned()).unwrap_or_else(|| v.to_string()))
    });
    xml
}

pub fn write_blocks(blocks: &[Block], sink: &mut dyn PartSink, out: &mut Vec<Node>) {
    for b in blocks {
        out.push(Node::Element(write_block(b, sink)));
    }
}

fn write_sdt<T>(s: &Sdt<T>, sink: &mut dyn PartSink, f: &mut dyn FnMut(&T, &mut dyn PartSink) -> Element) -> Element {
    let mut e = Element::new("w:sdt");
    if let Some(p) = &s.props {
        e.children.push(Node::Element(p.clone()));
    }
    if let Some(p) = &s.end_props {
        e.children.push(Node::Element(p.clone()));
    }
    let mut c = Element::new("w:sdtContent");
    for x in &s.content {
        c.children.push(Node::Element(f(x, sink)));
    }
    e.children.push(Node::Element(c));
    e
}

fn write_wrapper<T>(w: &Wrapper<T>, sink: &mut dyn PartSink, f: &mut dyn FnMut(&T, &mut dyn PartSink) -> Element) -> Element {
    let mut e = el(&w.name, &w.attrs);
    if let Some(p) = &w.props {
        e.children.push(Node::Element(p.clone()));
    }
    for x in &w.content {
        e.children.push(Node::Element(f(x, sink)));
    }
    e
}

fn write_revision<T>(r: &Revision<T>, sink: &mut dyn PartSink, f: &mut dyn FnMut(&T, &mut dyn PartSink) -> Element) -> Element {
    let mut e = el(&r.name, &r.attrs);
    for x in &r.content {
        e.children.push(Node::Element(f(x, sink)));
    }
    e
}

pub fn write_block(b: &Block, sink: &mut dyn PartSink) -> Element {
    match b {
        Block::Paragraph(p) => write_paragraph(p, sink),
        Block::Table(t) => write_table(t, sink),
        Block::Sdt(s) => write_sdt(s, sink, &mut |x, k| write_block(x, k)),
        Block::CustomXml(w) => write_wrapper(w, sink, &mut |x, k| write_block(x, k)),
        Block::Revision(r) => write_revision(r, sink, &mut |x, k| write_block(x, k)),
        Block::Marker(m) => el(&m.name, &m.attrs),
        Block::AltChunk(o) => write_object(o, sink),
    }
}

pub fn write_paragraph(p: &Paragraph, sink: &mut dyn PartSink) -> Element {
    let mut e = el("w:p", &p.attrs);
    if let Some(props) = &p.props {
        if !props.is_empty() {
            let mut pe = props.element().clone();
            write_sect_refs(&mut pe, sink);
            e.children.push(Node::Element(pe));
        }
    }
    for i in &p.content {
        e.children.push(Node::Element(write_inline(i, sink)));
    }
    e
}

/// A section break inside a paragraph names its headers and footers by part
/// index; the package writer swaps those for relationship ids. Nothing to do
/// at this level, but the hook keeps the two writers symmetrical.
fn write_sect_refs(_ppr: &mut Element, _sink: &mut dyn PartSink) {}

pub fn write_inline(i: &Inline, sink: &mut dyn PartSink) -> Element {
    match i {
        Inline::Run(r) => write_run(r, sink),
        Inline::Hyperlink(h) => {
            let mut e = Element::new("w:hyperlink");
            if let Some(t) = &h.target {
                let rid = sink.add(&Part::External {
                    rel_type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink".into(),
                    target: t.clone(),
                });
                e.attrs.push(("r:id".into(), rid));
            }
            e.attrs.extend(h.attrs.iter().cloned());
            for x in &h.content {
                e.children.push(Node::Element(write_inline(x, sink)));
            }
            e
        }
        Inline::SimpleField(f) => {
            let mut e = el("w:fldSimple", &f.attrs);
            if let Some(d) = &f.data {
                e.children.push(Node::Element(d.clone()));
            }
            for x in &f.content {
                e.children.push(Node::Element(write_inline(x, sink)));
            }
            e
        }
        Inline::Sdt(s) => write_sdt(s, sink, &mut |x, k| write_inline(x, k)),
        Inline::SmartTag(w) | Inline::CustomXml(w) | Inline::Bidi(w) => write_wrapper(w, sink, &mut |x, k| write_inline(x, k)),
        Inline::Revision(r) => write_revision(r, sink, &mut |x, k| write_inline(x, k)),
        Inline::Marker(m) => el(&m.name, &m.attrs),
        Inline::Math(o) => write_object(o, sink),
        Inline::Loose(item) => {
            let run = write_run(&Run { content: vec![item.clone()], ..Default::default() }, sink);
            match run.children.into_iter().next() {
                Some(Node::Element(e)) => e,
                _ => Element::new("w:r"),
            }
        }
    }
}

fn text_el(name: &str, t: &str) -> Element {
    let mut e = Element::new(name);
    if t.starts_with(' ') || t.ends_with(' ') || t.contains("  ") || t.contains('\t') || t.contains('\n') {
        e.attrs.push(("xml:space".into(), "preserve".into()));
    }
    if !t.is_empty() {
        e.children.push(Node::Text(t.to_string()));
    }
    e
}

pub fn write_run(r: &Run, sink: &mut dyn PartSink) -> Element {
    let mut e = Element::new("w:r");
    if let Some(p) = &r.props {
        if !p.is_empty() {
            e.children.push(Node::Element(p.element().clone()));
        }
    }
    for item in &r.content {
        let c = match item {
            RunItem::Text(t) => text_el("w:t", t),
            RunItem::DeletedText(t) => text_el("w:delText", t),
            RunItem::InstrText(t) => text_el("w:instrText", t),
            RunItem::DeletedInstrText(t) => text_el("w:delInstrText", t),
            RunItem::Tab => Element::new("w:tab"),
            RunItem::Break { attrs } => el("w:br", attrs),
            RunItem::FieldChar { attrs, data } => {
                let mut f = el("w:fldChar", attrs);
                if let Some(d) = data {
                    f.children.push(Node::Element(d.clone()));
                }
                f
            }
            RunItem::Mark { name, attrs } => el(name, attrs),
            RunItem::Ruby(ruby) => {
                let mut rb = Element::new("w:ruby");
                rb.children.push(Node::Element(ruby.props.clone()));
                let mut rt = Element::new("w:rt");
                for x in &ruby.text {
                    rt.children.push(Node::Element(write_inline(x, sink)));
                }
                let mut base = Element::new("w:rubyBase");
                for x in &ruby.base {
                    base.children.push(Node::Element(write_inline(x, sink)));
                }
                rb.children.push(Node::Element(rt));
                rb.children.push(Node::Element(base));
                rb
            }
            RunItem::Object(o) => write_object(o, sink),
            RunItem::Props(p) => p.element().clone(),
        };
        e.children.push(Node::Element(c));
    }
    e
}

fn write_table(t: &Table, sink: &mut dyn PartSink) -> Element {
    let mut e = Element::new("w:tbl");
    e.children.push(Node::Element(t.props.element().clone()));
    if let Some(g) = &t.grid {
        e.children.push(Node::Element(g.clone()));
    }
    for r in &t.rows {
        e.children.push(Node::Element(write_row_item(r, sink)));
    }
    e
}

fn write_row_item(r: &RowItem, sink: &mut dyn PartSink) -> Element {
    match r {
        RowItem::Row(row) => {
            let mut e = el("w:tr", &row.attrs);
            if let Some(p) = &row.props_ex {
                e.children.push(Node::Element(p.element().clone()));
            }
            let props = row.props.as_ref().filter(|p| !p.is_empty()).map(|p| Node::Element(p.element().clone()));
            if !row.props_after_cells {
                e.children.extend(props.clone());
            }
            for c in &row.cells {
                e.children.push(Node::Element(write_cell_item(c, sink)));
            }
            if row.props_after_cells {
                e.children.extend(props);
            }
            e
        }
        RowItem::Sdt(s) => write_sdt(s, sink, &mut |x, k| write_row_item(x, k)),
        RowItem::CustomXml(w) => write_wrapper(w, sink, &mut |x, k| write_row_item(x, k)),
        RowItem::Revision(rv) => write_revision(rv, sink, &mut |x, k| write_row_item(x, k)),
        RowItem::Marker(m) => el(&m.name, &m.attrs),
    }
}

fn write_cell_item(c: &CellItem, sink: &mut dyn PartSink) -> Element {
    match c {
        CellItem::Cell(cell) => {
            let mut e = el("w:tc", &cell.attrs);
            if let Some(p) = &cell.props {
                if !p.is_empty() {
                    e.children.push(Node::Element(p.element().clone()));
                }
            }
            for b in &cell.content {
                e.children.push(Node::Element(write_block(b, sink)));
            }
            e
        }
        CellItem::Sdt(s) => write_sdt(s, sink, &mut |x, k| write_cell_item(x, k)),
        CellItem::CustomXml(w) => write_wrapper(w, sink, &mut |x, k| write_cell_item(x, k)),
        CellItem::Revision(rv) => write_revision(rv, sink, &mut |x, k| write_cell_item(x, k)),
        CellItem::Marker(m) => el(&m.name, &m.attrs),
    }
}
