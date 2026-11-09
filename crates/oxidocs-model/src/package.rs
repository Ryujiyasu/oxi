// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! The .docx package: reading a file into a [`DocxModel`] and writing a model
//! back out as a file, either on its own or over a template that supplies
//! the document-wide parts (styles, numbering, settings, theme, fonts).

use std::collections::{BTreeMap, HashMap};
use std::io::{Cursor, Read, Write};

use serde::{Deserialize, Serialize};

use crate::content::{self, Block, Object, Part, PartSink, Resolver};
use crate::docparts::{self, Numbering, Styles};
use crate::props::{Props, Unsupported};
use crate::xml::{self, Element, Node};

pub const REL_NS: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
pub const REL_OFFICE_DOC: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
pub const REL_HEADER: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header";
pub const REL_FOOTER: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer";
pub const REL_FOOTNOTES: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes";
pub const REL_ENDNOTES: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes";

const CT_DOCUMENT: &str = "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml";
const CT_HEADER: &str = "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml";
const CT_FOOTER: &str = "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml";
const CT_FOOTNOTES: &str = "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml";
const CT_ENDNOTES: &str = "application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml";
const CT_RELS: &str = "application/vnd.openxmlformats-package.relationships+xml";

const REL_BASE: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/";
const REL_SWE: &str = "http://schemas.microsoft.com/office/2007/relationships/stylesWithEffects";
const REL_PEOPLE: &str = "http://schemas.microsoft.com/office/2011/relationships/people";
const CT_WML: &str = "application/vnd.openxmlformats-officedocument.wordprocessingml.";

/// Relationship types that point at the document's content. Everything else
/// a template's document links to (styles, numbering, settings, theme,
/// fonts, comments, custom XML ...) is document-wide and is carried over.
const CONTENT_RELS: &[&str] = &[
    "/header", "/footer", "/footnotes", "/endnotes", "/image", "/hyperlink",
    "/oleObject", "/package", "/chart", "/diagramData", "/diagramLayout",
    "/diagramQuickStyle", "/diagramColors", "/diagramDrawing", "/aFChunk",
    "/customXmlProps", "/video", "/audio", "/media", "/control",
    "/chartEx", "/ink", "/model3d", "/hdphoto", "/svgImage", "/subDocument",
];

pub fn is_content_rel(rel_type: &str) -> bool {
    CONTENT_RELS.iter().any(|s| rel_type.ends_with(s)) && !rel_type.ends_with("/customXml")
}

// ---------------------------------------------------------------------------
// Package files
// ---------------------------------------------------------------------------

#[derive(Debug, Clone)]
pub struct Rel {
    pub id: String,
    pub rel_type: String,
    pub target: String,
    pub external: bool,
}

/// A zip package held in memory.
#[derive(Debug, Clone, Default)]
pub struct Package {
    pub files: BTreeMap<String, Vec<u8>>,
}

#[derive(Debug)]
pub enum Error {
    Zip(String),
    Xml(String),
    Missing(String),
    Unsupported(Unsupported),
}

impl std::fmt::Display for Error {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Error::Zip(s) => write!(f, "zip: {s}"),
            Error::Xml(s) => write!(f, "xml: {s}"),
            Error::Missing(s) => write!(f, "missing part: {s}"),
            Error::Unsupported(u) => write!(f, "{u}"),
        }
    }
}

impl std::error::Error for Error {}

impl From<Unsupported> for Error {
    fn from(u: Unsupported) -> Self {
        Error::Unsupported(u)
    }
}

impl From<xml::XmlError> for Error {
    fn from(e: xml::XmlError) -> Self {
        Error::Xml(e.0)
    }
}

impl Package {
    pub fn read(bytes: &[u8]) -> Result<Package, Error> {
        let mut zip = zip::ZipArchive::new(Cursor::new(bytes)).map_err(|e| Error::Zip(e.to_string()))?;
        let mut files = BTreeMap::new();
        for i in 0..zip.len() {
            let mut f = zip.by_index(i).map_err(|e| Error::Zip(e.to_string()))?;
            if f.is_dir() {
                continue;
            }
            let mut buf = Vec::new();
            f.read_to_end(&mut buf).map_err(|e| Error::Zip(e.to_string()))?;
            files.insert(f.name().trim_start_matches('/').to_string(), buf);
        }
        Ok(Package { files })
    }

    pub fn write(&self) -> Vec<u8> {
        let mut zip = zip::ZipWriter::new(Cursor::new(Vec::new()));
        let opts = zip::write::SimpleFileOptions::default().compression_method(zip::CompressionMethod::Deflated);
        // [Content_Types].xml first, as Word writes it.
        let mut names: Vec<&String> = self.files.keys().collect();
        names.sort_by_key(|n| (n.as_str() != "[Content_Types].xml", n.as_str() != "_rels/.rels", n.to_string()));
        for n in names {
            zip.start_file(n.as_str(), opts).expect("zip entry");
            zip.write_all(&self.files[n]).expect("zip write");
        }
        zip.finish().expect("zip finish").into_inner()
    }

    pub fn xml(&self, name: &str) -> Result<Element, Error> {
        let bytes = self.files.get(name).ok_or_else(|| Error::Missing(name.to_string()))?;
        let text = decode(bytes);
        Ok(xml::parse(&text)?)
    }

    pub fn rels(&self, part: &str) -> Vec<Rel> {
        let path = rels_path(part);
        let Ok(root) = self.xml(&path) else { return Vec::new() };
        root.elements()
            .filter(|e| e.name.ends_with("Relationship"))
            .map(|e| {
                let external = e.attr("TargetMode") == Some("External");
                let target = e.attr("Target").unwrap_or("").to_string();
                Rel {
                    id: e.attr("Id").unwrap_or("").to_string(),
                    rel_type: e.attr("Type").unwrap_or("").to_string(),
                    target: if external { target } else { resolve(part, &target) },
                    external,
                }
            })
            .collect()
    }

    pub fn content_type(&self, name: &str) -> String {
        let Ok(ct) = self.xml("[Content_Types].xml") else { return String::new() };
        let slash = format!("/{name}");
        for e in ct.elements() {
            if e.name.ends_with("Override") && e.attr("PartName").map(|p| p.eq_ignore_ascii_case(&slash)) == Some(true) {
                return e.attr("ContentType").unwrap_or("").to_string();
            }
        }
        let ext = name.rsplit('.').next().unwrap_or("").to_ascii_lowercase();
        for e in ct.elements() {
            if e.name.ends_with("Default") && e.attr("Extension").map(|x| x.to_ascii_lowercase()) == Some(ext.clone()) {
                return e.attr("ContentType").unwrap_or("").to_string();
            }
        }
        String::new()
    }

    /// The main document part's name.
    pub fn main_part(&self) -> Result<String, Error> {
        self.rels("")
            .into_iter()
            .find(|r| r.rel_type == REL_OFFICE_DOC || r.rel_type.ends_with("/officeDocument"))
            .map(|r| r.target)
            .ok_or_else(|| Error::Missing("officeDocument relationship".into()))
    }
}

/// Text of an XML part, whatever its encoding declaration (UTF-8 or UTF-16).
pub fn decode(bytes: &[u8]) -> String {
    if bytes.starts_with(&[0xFF, 0xFE]) || bytes.starts_with(&[0xFE, 0xFF]) {
        let le = bytes[0] == 0xFF;
        let units: Vec<u16> = bytes[2..]
            .chunks_exact(2)
            .map(|c| if le { u16::from_le_bytes([c[0], c[1]]) } else { u16::from_be_bytes([c[0], c[1]]) })
            .collect();
        return String::from_utf16_lossy(&units);
    }
    let b = bytes.strip_prefix(&[0xEF, 0xBB, 0xBF]).unwrap_or(bytes);
    String::from_utf8_lossy(b).into_owned()
}

pub fn rels_path(part: &str) -> String {
    match part.rsplit_once('/') {
        Some((dir, file)) => format!("{dir}/_rels/{file}.rels"),
        None if part.is_empty() => "_rels/.rels".to_string(),
        None => format!("_rels/{part}.rels"),
    }
}

/// Resolve a relationship target against the part that holds it.
pub fn resolve(source: &str, target: &str) -> String {
    let target = target.replace('\\', "/");
    let mut parts: Vec<&str> = if let Some(abs) = target.strip_prefix('/') {
        return normalize(abs.split('/').collect());
    } else {
        source.rsplit_once('/').map(|(d, _)| d.split('/').collect()).unwrap_or_default()
    };
    parts.extend(target.split('/'));
    normalize(parts)
}

fn normalize(parts: Vec<&str>) -> String {
    let mut out: Vec<&str> = Vec::new();
    for p in parts {
        match p {
            "" | "." => {}
            ".." => {
                out.pop();
            }
            p => out.push(p),
        }
    }
    out.join("/")
}

/// The path of `target` relative to the directory of `source`.
pub fn relative(source: &str, target: &str) -> String {
    let src: Vec<&str> = source.rsplit_once('/').map(|(d, _)| d.split('/').collect()).unwrap_or_default();
    let dst: Vec<&str> = target.split('/').collect();
    let common = src.iter().zip(dst.iter()).take_while(|(a, b)| a == b).count();
    let mut out: Vec<String> = std::iter::repeat_n("..".to_string(), src.len() - common).collect();
    out.extend(dst[common..].iter().map(|s| s.to_string()));
    out.join("/")
}

// ---------------------------------------------------------------------------
// Model
// ---------------------------------------------------------------------------

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct HeaderFooter {
    /// "w:hdr" or "w:ftr".
    pub kind: String,
    pub content: Vec<Block>,
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct Note {
    /// w:id, w:type (separator, continuationSeparator, ...).
    pub attrs: Vec<(String, String)>,
    pub content: Vec<Block>,
}

/// A whole document as data.
#[derive(Debug, Clone, PartialEq, Default, Serialize, Deserialize)]
pub struct DocxModel {
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub background: Option<Element>,
    pub body: Vec<Block>,
    /// The final section's properties. Header and footer references name a
    /// part by index into `headers_footers` (`r:id="#3"`).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub section: Option<Props>,
    /// Where the body's own sectPr sits when it is not the last thing in
    /// the body (a few generators write content after it): its index among
    /// the blocks. None means last, as the schema requires.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub section_at: Option<usize>,
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub headers_footers: Vec<HeaderFooter>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub footnotes: Option<Vec<Note>>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub endnotes: Option<Vec<Note>>,

    // Document-wide parts. A writer given a template takes any of these the
    // model leaves empty from the template.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub styles: Option<Styles>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub styles_with_effects: Option<Styles>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub numbering: Option<Numbering>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub settings: Option<Object>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub web_settings: Option<Object>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub font_table: Option<Object>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub people: Option<Object>,
    /// The theme part (colours, fonts, effects), applied whole.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub theme: Option<Part>,
    /// Custom XML data stores (content controls bind to them).
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub custom_xml: Vec<Part>,
    /// The building-blocks (glossary) document, applied whole.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub glossary: Option<Part>,
    /// Data parts Word keeps for its own features, carried whole: keyboard
    /// customizations, Office intelligence (editor suggestions) state,
    /// document tasks.
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub app_data: Vec<Part>,
    /// Comments: each a w:comment with its id, author, date and content.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub comments: Option<Vec<Note>>,
    /// Comment threads and resolved state (w15:commentsEx), durable ids
    /// (w16cid:commentsIds) and extensible data (w16cex:commentsExtensible).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub comments_extended: Option<Object>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub comments_ids: Option<Object>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub comments_extensible: Option<Object>,
}

/// Which document-wide part a relationship type names.
fn doc_part_kind(rel_type: &str) -> Option<&'static str> {
    let tail = rel_type.rsplit('/').next().unwrap_or("");
    Some(match tail {
        "styles" => "styles",
        "stylesWithEffects" => "stylesWithEffects",
        "numbering" => "numbering",
        "settings" => "settings",
        "webSettings" => "webSettings",
        "fontTable" => "fontTable",
        "people" => "people",
        "theme" => "theme",
        "customXml" => "customXml",
        "glossaryDocument" => "glossaryDocument",
        "keyMapCustomizations" | "intelligence" | "documenttasks" => "appData",
        "comments" => "comments",
        "commentsExtended" => "commentsExtended",
        "commentsIds" => "commentsIds",
        "commentsExtensible" => "commentsExtensible",
        _ => return None,
    })
}

struct PartResolver<'a> {
    pkg: &'a Package,
    part: String,
    rels: HashMap<String, Rel>,
    /// Header and footer parts already read, by part name, so sections that
    /// share one header name the same entry.
    depth: usize,
}

impl<'a> PartResolver<'a> {
    fn new(pkg: &'a Package, part: &str) -> Self {
        let rels = pkg.rels(part).into_iter().map(|r| (r.id.clone(), r)).collect();
        PartResolver { pkg, part: part.to_string(), rels, depth: 0 }
    }
}

fn load_part(pkg: &Package, rel: &Rel, depth: usize) -> Result<Part, Unsupported> {
    if rel.external {
        return Ok(Part::External { rel_type: rel.rel_type.clone(), target: rel.target.clone() });
    }
    let Some(data) = pkg.files.get(&rel.target) else {
        // A dangling internal link: keep it as an external target so the
        // reference survives as written.
        return Ok(Part::External { rel_type: rel.rel_type.clone(), target: format!("missing:{}", rel.target) });
    };
    let mut rels = Vec::new();
    if depth < 4 {
        for r in pkg.rels(&rel.target) {
            rels.push((r.id.clone(), load_part(pkg, &r, depth + 1)?));
        }
    }
    let name = relative("word/document.xml", &rel.target);
    Ok(Part::Internal {
        rel_type: rel.rel_type.clone(),
        name,
        content_type: pkg.content_type(&rel.target),
        data: data.clone(),
        rels,
    })
}

impl Resolver for PartResolver<'_> {
    fn part(&mut self, rid: &str) -> Result<Part, Unsupported> {
        let rel = self
            .rels
            .get(rid)
            .ok_or_else(|| Unsupported { path: format!("{}#{rid}(dangling)", self.part) })?;
        load_part(self.pkg, rel, self.depth)
    }
}

impl DocxModel {
    /// Whether the model itself carries a document-wide part of this kind.
    fn provides(&self, kind: &str) -> bool {
        match kind {
            "styles" => self.styles.is_some(),
            "stylesWithEffects" => self.styles_with_effects.is_some(),
            "numbering" => self.numbering.is_some(),
            "settings" => self.settings.is_some(),
            "webSettings" => self.web_settings.is_some(),
            "fontTable" => self.font_table.is_some(),
            "people" => self.people.is_some(),
            "theme" => self.theme.is_some(),
            "customXml" => !self.custom_xml.is_empty(),
            "glossaryDocument" => self.glossary.is_some(),
            "appData" => !self.app_data.is_empty(),
            "comments" => self.comments.is_some(),
            "commentsExtended" => self.comments_extended.is_some(),
            "commentsIds" => self.comments_ids.is_some(),
            "commentsExtensible" => self.comments_extensible.is_some(),
            _ => false,
        }
    }

    /// The same document without its document-wide parts, for writing over
    /// a template that supplies them.
    pub fn without_document_parts(mut self) -> DocxModel {
        self.styles = None;
        self.styles_with_effects = None;
        self.numbering = None;
        self.settings = None;
        self.web_settings = None;
        self.font_table = None;
        self.people = None;
        self.theme = None;
        self.custom_xml.clear();
        self.glossary = None;
        self.app_data.clear();
        self.comments = None;
        self.comments_extended = None;
        self.comments_ids = None;
        self.comments_extensible = None;
        self
    }

    pub fn read(bytes: &[u8]) -> Result<DocxModel, Error> {
        let pkg = Package::read(bytes)?;
        DocxModel::from_package(&pkg)
    }

    pub fn from_package(pkg: &Package) -> Result<DocxModel, Error> {
        let main = pkg.main_part()?;
        let root = pkg.xml(&main)?;
        let mut res = PartResolver::new(pkg, &main);
        let mut model = DocxModel::default();
        let mut hf_index: HashMap<String, usize> = HashMap::new();
        for c in root.elements() {
            match c.name.as_str() {
                "w:background" => model.background = Some(c.clone()),
                "w:body" => {
                    let mut blocks: Vec<&Element> = c.elements().collect();
                    let sects: Vec<usize> = blocks.iter().enumerate().filter(|(_, e)| e.name == "w:sectPr").map(|(i, _)| i).collect();
                    if sects.len() > 1 {
                        return Err(Error::Unsupported(Unsupported { path: "w:body/w:sectPr(several)".into() }));
                    }
                    if let Some(&at) = sects.first() {
                        let s = blocks.remove(at);
                        let mut props = Props::read("w:body", s)?;
                        link_sections(&mut props, pkg, &res.rels, &mut model, &mut hf_index)?;
                        model.section = Some(props);
                        if at != blocks.len() {
                            // Count the blocks before it as the model will hold them.
                            let mut before = Vec::new();
                            for b in &blocks[..at] {
                                let one = content::read_blocks(
                                    &Element { name: "w:body".into(), attrs: vec![], children: vec![Node::Element((*b).clone())] },
                                    "w:body",
                                    &mut res,
                                )?;
                                before.extend(one);
                            }
                            model.section_at = Some(before.len());
                        }
                    }
                    let mut body = Vec::new();
                    for b in blocks {
                        let one = content::read_blocks(
                            &Element { name: "w:body".into(), attrs: vec![], children: vec![Node::Element(b.clone())] },
                            "w:body",
                            &mut res,
                        )?;
                        body.extend(one);
                    }
                    for_each_paragraph_sect(&mut body, &mut |props| {
                        link_sections(props, pkg, &res.rels, &mut model, &mut hf_index)
                    })?;
                    model.body = body;
                }
                other => return Err(Error::Unsupported(Unsupported { path: format!("w:document/{other}") })),
            }
        }
        for rel in pkg.rels(&main) {
            if rel.external || is_content_rel(&rel.rel_type) {
                continue;
            }
            // A link to a part the package does not hold says nothing.
            if !pkg.files.contains_key(&rel.target) {
                continue;
            }
            let Some(kind) = doc_part_kind(&rel.rel_type) else {
                return Err(Error::Unsupported(Unsupported { path: format!("document part {}", rel.rel_type) }));
            };
            let mut pres = PartResolver::new(pkg, &rel.target);
            match kind {
                "styles" | "stylesWithEffects" => {
                    let st = Styles::read(&pkg.xml(&rel.target)?)?;
                    if kind == "styles" {
                        model.styles = Some(st);
                    } else {
                        model.styles_with_effects = Some(st);
                    }
                }
                "numbering" => model.numbering = Some(Numbering::read(&pkg.xml(&rel.target)?, &mut pres)?),
                "settings" => {
                    model.settings = Some(docparts::read_settings_bag(&pkg.xml(&rel.target)?, "w:settings", docparts::SETTINGS, &mut pres)?)
                }
                "webSettings" => {
                    model.web_settings = Some(docparts::read_settings_bag(&pkg.xml(&rel.target)?, "w:webSettings", docparts::WEB_SETTINGS, &mut pres)?)
                }
                "fontTable" => model.font_table = Some(docparts::read_font_table(&pkg.xml(&rel.target)?, &mut pres)?),
                "people" => model.people = Some(docparts::read_people(&pkg.xml(&rel.target)?, &mut pres)?),
                "theme" => model.theme = Some(load_part(pkg, &rel, 0)?),
                "customXml" => model.custom_xml.push(load_part(pkg, &rel, 0)?),
                "glossaryDocument" => model.glossary = Some(load_part(pkg, &rel, 0)?),
                "appData" => model.app_data.push(load_part(pkg, &rel, 0)?),
                "comments" => {
                    let root = pkg.xml(&rel.target)?;
                    if root.name != "w:comments" {
                        return Err(Error::Unsupported(Unsupported { path: format!("comments/{}", root.name) }));
                    }
                    let mut out = Vec::new();
                    for c in root.elements() {
                        if c.name != "w:comment" {
                            return Err(Error::Unsupported(Unsupported { path: format!("w:comments/{}", c.name) }));
                        }
                        let attrs = c.attrs.iter().filter(|(k, _)| !crate::props::is_session_attr(k)).cloned().collect();
                        out.push(Note { attrs, content: content::read_blocks(c, "w:comment", &mut pres)? });
                    }
                    model.comments = Some(out);
                }
                "commentsExtended" => {
                    model.comments_extended = Some(docparts::read_list(&pkg.xml(&rel.target)?, "w15:commentsEx", &["w15:commentEx"], &mut pres)?)
                }
                "commentsIds" => {
                    model.comments_ids = Some(docparts::read_list(&pkg.xml(&rel.target)?, "w16cid:commentsIds", &["w16cid:commentId"], &mut pres)?)
                }
                "commentsExtensible" => {
                    model.comments_extensible = Some(docparts::read_list(
                        &pkg.xml(&rel.target)?,
                        "w16cex:commentsExtensible",
                        &["w16cex:commentExtensible", "w16cex:extLst"],
                        &mut pres,
                    )?)
                }
                _ => unreachable!(),
            }
        }
        for rel in pkg.rels(&main) {
            if rel.rel_type == REL_FOOTNOTES || rel.rel_type == REL_ENDNOTES {
                let notes = read_notes(pkg, &rel.target)?;
                if rel.rel_type == REL_FOOTNOTES {
                    model.footnotes = Some(notes);
                } else {
                    model.endnotes = Some(notes);
                }
            }
        }
        Ok(model)
    }
}

fn read_notes(pkg: &Package, part: &str) -> Result<Vec<Note>, Error> {
    let root = pkg.xml(part)?;
    let mut res = PartResolver::new(pkg, part);
    let mut out = Vec::new();
    for n in root.elements() {
        if n.name != "w:footnote" && n.name != "w:endnote" {
            return Err(Error::Unsupported(Unsupported { path: format!("{}/{}", root.name, n.name) }));
        }
        let attrs = n.attrs.iter().filter(|(k, _)| !crate::props::is_session_attr(k)).cloned().collect();
        out.push(Note { attrs, content: content::read_blocks(n, &n.name, &mut res)? });
    }
    Ok(out)
}

/// Visit the section properties carried by paragraphs (section breaks).
pub fn for_each_paragraph_sect(
    blocks: &mut [Block],
    f: &mut dyn FnMut(&mut Props) -> Result<(), Error>,
) -> Result<(), Error> {
    for b in blocks.iter_mut() {
        match b {
            Block::Paragraph(p) => {
                if let Some(props) = p.props.as_mut() {
                    if let Some(s) = props.get("w:sectPr").cloned() {
                        let mut sp = Props::read("w:pPr", &s)?;
                        f(&mut sp)?;
                        props.set(sp.into_element()).map_err(Error::Unsupported)?;
                    }
                }
            }
            Block::Sdt(s) => for_each_paragraph_sect(&mut s.content, f)?,
            Block::CustomXml(w) => for_each_paragraph_sect(&mut w.content, f)?,
            Block::Revision(r) => for_each_paragraph_sect(&mut r.content, f)?,
            _ => {}
        }
    }
    Ok(())
}

/// Swap header/footer relationship ids for indexes into the model's list.
fn link_sections(
    props: &mut Props,
    pkg: &Package,
    rels: &HashMap<String, Rel>,
    model: &mut DocxModel,
    seen: &mut HashMap<String, usize>,
) -> Result<(), Error> {
    let el = props.element_mut();
    for c in el.elements_mut() {
        if c.name != "w:headerReference" && c.name != "w:footerReference" {
            continue;
        }
        let Some(rid) = c.attr("r:id").map(|s| s.to_string()) else { continue };
        let Some(rel) = rels.get(&rid) else {
            // The reference names a relationship the document lacks: it
            // points nowhere, and so does the copy.
            c.set_attr("r:id", content::dangling(&rid));
            continue;
        };
        let index = match seen.get(&rel.target) {
            Some(i) => *i,
            None => {
                let root = pkg.xml(&rel.target)?;
                let mut res = PartResolver::new(pkg, &rel.target);
                let content = content::read_blocks(&root, &root.name, &mut res)?;
                model.headers_footers.push(HeaderFooter { kind: root.name.clone(), content });
                let i = model.headers_footers.len() - 1;
                seen.insert(rel.target.clone(), i);
                i
            }
        };
        c.set_attr("r:id", format!("#{index}"));
    }
    Ok(())
}

// ---------------------------------------------------------------------------
// Writing
// ---------------------------------------------------------------------------

struct OutPackage {
    files: BTreeMap<String, Vec<u8>>,
    types: BTreeMap<String, String>,
    /// Internal parts already written, so a picture used twice is stored once.
    written: Vec<(Part, String)>,
}

impl OutPackage {
    fn unique_name(&self, hint: &str) -> String {
        let hint = resolve("word/document.xml", hint.trim_start_matches('/'));
        if !self.files.contains_key(&hint) {
            return hint;
        }
        let (stem, ext) = match hint.rsplit_once('.') {
            Some((s, e)) if !s.ends_with('/') => (s.to_string(), format!(".{e}")),
            _ => (hint.clone(), String::new()),
        };
        (1..).map(|n| format!("{stem}_{n}{ext}")).find(|n| !self.files.contains_key(n)).unwrap()
    }

    fn put(&mut self, name: &str, content_type: &str, data: Vec<u8>) {
        if !content_type.is_empty() {
            self.types.insert(name.to_string(), content_type.to_string());
        }
        self.files.insert(name.to_string(), data);
    }

    /// Write an internal part (and the parts it links to); return its name.
    fn add_internal(&mut self, part: &Part) -> String {
        if let Some((_, n)) = self.written.iter().find(|(p, _)| p == part) {
            return n.clone();
        }
        let Part::Internal { name, content_type, data, rels, .. } = part else { unreachable!() };
        let path = self.unique_name(name);
        self.put(&path, content_type, data.clone());
        self.written.push((part.clone(), path.clone()));
        if !rels.is_empty() {
            let mut out = Vec::new();
            for (rid, p) in rels {
                match p {
                    Part::External { rel_type, target } => out.push(rel_xml(rid, rel_type, target, true)),
                    Part::Internal { rel_type, .. } => {
                        let n = self.add_internal(p);
                        out.push(rel_xml(rid, rel_type, &relative(&path, &n), false));
                    }
                }
            }
            self.files.insert(rels_path(&path), rels_doc(&out));
        }
        path
    }
}

fn rel_xml(id: &str, rel_type: &str, target: &str, external: bool) -> String {
    let mode = if external { r#" TargetMode="External""# } else { "" };
    format!(
        r#"<Relationship Id="{}" Type="{}" Target="{}"{mode}/>"#,
        xml::escape(id, true),
        xml::escape(rel_type, true),
        xml::escape(target, true)
    )
}

fn rels_doc(rels: &[String]) -> Vec<u8> {
    format!(
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">{}</Relationships>",
        rels.concat()
    )
    .into_bytes()
}

/// Relationship ids for one source part.
struct PartRels<'a> {
    pkg: &'a mut OutPackage,
    source: String,
    rels: Vec<String>,
    next: usize,
    taken: Vec<String>,
}

impl<'a> PartRels<'a> {
    fn new(pkg: &'a mut OutPackage, source: &str, taken: Vec<String>) -> Self {
        PartRels { pkg, source: source.to_string(), rels: Vec::new(), next: 1, taken }
    }

    fn fresh_id(&mut self) -> String {
        loop {
            let id = format!("rId{}", self.next);
            self.next += 1;
            if !self.taken.contains(&id) {
                self.taken.push(id.clone());
                return id;
            }
        }
    }

    fn link(&mut self, rel_type: &str, target_part: &str) -> String {
        let id = self.fresh_id();
        let rel = rel_xml(&id, rel_type, &relative(&self.source, target_part), false);
        self.rels.push(rel);
        id
    }
}

impl PartSink for PartRels<'_> {
    fn add(&mut self, part: &Part) -> String {
        let id = self.fresh_id();
        match part {
            Part::External { rel_type, target } => {
                let target = target.strip_prefix("missing:").map(|t| relative(&self.source, t)).unwrap_or_else(|| target.clone());
                let external = !part_is_missing(part);
                self.rels.push(rel_xml(&id, rel_type, &target, external));
            }
            Part::Internal { rel_type, .. } => {
                let name = self.pkg.add_internal(part);
                self.rels.push(rel_xml(&id, rel_type, &relative(&self.source, &name), false));
            }
        }
        id
    }
}

fn part_is_missing(p: &Part) -> bool {
    matches!(p, Part::External { target, .. } if target.starts_with("missing:"))
}

fn root_with(name: &str, children: Vec<Node>) -> Element {
    Element { name: name.to_string(), attrs: Vec::new(), children }
}

/// Copy a part and everything it links to (transitively) from a package.
fn copy_tree(src: &Package, part: &str, out: &mut OutPackage, done: &mut Vec<String>) {
    if done.iter().any(|d| d == part) {
        return;
    }
    done.push(part.to_string());
    if let Some(data) = src.files.get(part) {
        out.put(part, &src.content_type(part), data.clone());
    }
    let rp = rels_path(part);
    if let Some(data) = src.files.get(&rp) {
        out.files.insert(rp, data.clone());
    }
    for r in src.rels(part) {
        if !r.external {
            copy_tree(src, &r.target, out, done);
        }
    }
}

impl DocxModel {
    /// Write the model as a .docx. With a template, the template's
    /// document-wide parts (styles, numbering, settings, theme, fonts,
    /// comments, package properties) are carried over; the content parts are
    /// all written from the model.
    pub fn write(&self, template: Option<&Package>) -> Result<Vec<u8>, Error> {
        let mut out = OutPackage { files: BTreeMap::new(), types: BTreeMap::new(), written: Vec::new() };
        let main = "word/document.xml".to_string();
        let mut doc_rels: Vec<String> = Vec::new();
        let mut taken: Vec<String> = Vec::new();
        let mut done = vec![main.clone()];

        if let Some(t) = template {
            let t_main = t.main_part()?;
            // Package-level parts (docProps, thumbnail ...) except the main
            // document, which we write.
            let root_rels: Vec<Rel> = t.rels("");
            let mut root_out = Vec::new();
            for r in &root_rels {
                if r.rel_type == REL_OFFICE_DOC || r.rel_type.ends_with("/officeDocument") {
                    root_out.push(rel_xml(&r.id, &r.rel_type, &main, false));
                } else if r.external {
                    root_out.push(rel_xml(&r.id, &r.rel_type, &r.target, true));
                } else {
                    copy_tree(t, &r.target, &mut out, &mut done);
                    root_out.push(rel_xml(&r.id, &r.rel_type, &r.target, false));
                }
            }
            out.files.insert("_rels/.rels".into(), rels_doc(&root_out));
            for r in t.rels(&t_main) {
                if is_content_rel(&r.rel_type) || doc_part_kind(&r.rel_type).map(|k| self.provides(k)) == Some(true) {
                    continue;
                }
                taken.push(r.id.clone());
                if r.external {
                    doc_rels.push(rel_xml(&r.id, &r.rel_type, &r.target, true));
                } else {
                    let target = if r.target == t_main { main.clone() } else { r.target.clone() };
                    copy_tree(t, &target, &mut out, &mut done);
                    doc_rels.push(rel_xml(&r.id, &r.rel_type, &relative(&main, &target), false));
                }
            }
        } else {
            out.files.insert(
                "_rels/.rels".into(),
                rels_doc(&[rel_xml("rId1", REL_OFFICE_DOC, &main, false)]),
            );
        }

        // Headers and footers, each with its own relationships.
        let mut hf_names = Vec::new();
        for (i, hf) in self.headers_footers.iter().enumerate() {
            let header = hf.kind == "w:hdr";
            let name = out.unique_name(&format!("{}{}.xml", if header { "header" } else { "footer" }, i + 1));
            let mut sink = PartRels::new(&mut out, &name, Vec::new());
            let mut children = Vec::new();
            content::write_blocks(&hf.content, &mut sink, &mut children);
            let rels = std::mem::take(&mut sink.rels);
            drop(sink);
            let root = root_with(&hf.kind, children);
            out.put(&name, if header { CT_HEADER } else { CT_FOOTER }, xml::write(&root, true).into_bytes());
            if !rels.is_empty() {
                out.files.insert(rels_path(&name), rels_doc(&rels));
            }
            hf_names.push((name, header));
        }

        // Notes.
        let mut note_parts = Vec::new();
        for (notes, header, file, ct, rel) in [
            (&self.footnotes, "w:footnotes", "footnotes.xml", CT_FOOTNOTES, REL_FOOTNOTES),
            (&self.endnotes, "w:endnotes", "endnotes.xml", CT_ENDNOTES, REL_ENDNOTES),
        ] {
            let Some(notes) = notes else { continue };
            let name = out.unique_name(file);
            let mut sink = PartRels::new(&mut out, &name, Vec::new());
            let item = if header == "w:footnotes" { "w:footnote" } else { "w:endnote" };
            let mut children = Vec::new();
            for n in notes {
                let mut kids = Vec::new();
                content::write_blocks(&n.content, &mut sink, &mut kids);
                children.push(Node::Element(Element { name: item.into(), attrs: n.attrs.clone(), children: kids }));
            }
            let rels = std::mem::take(&mut sink.rels);
            drop(sink);
            out.put(&name, ct, xml::write(&root_with(header, children), true).into_bytes());
            if !rels.is_empty() {
                out.files.insert(rels_path(&name), rels_doc(&rels));
            }
            note_parts.push((name, rel));
        }

        // Document-wide parts the model carries.
        let mut doc_part_names: Vec<(String, String)> = Vec::new();
        {
            type Build<'b> = Box<dyn Fn(&mut dyn PartSink) -> Element + 'b>;
            let mut typed: Vec<(&str, &str, String, Build)> = Vec::new();
            if let Some(s) = &self.styles {
                typed.push(("styles.xml", "styles", format!("{CT_WML}styles+xml"), Box::new(move |_| s.write())));
            }
            if let Some(s) = &self.styles_with_effects {
                typed.push(("stylesWithEffects.xml", "stylesWithEffects", "application/vnd.ms-word.stylesWithEffects+xml".into(), Box::new(move |_| s.write())));
            }
            if let Some(n) = &self.numbering {
                typed.push(("numbering.xml", "numbering", format!("{CT_WML}numbering+xml"), Box::new(move |k| n.write(k))));
            }
            if let Some(list) = &self.comments {
                typed.push(("comments.xml", "comments", format!("{CT_WML}comments+xml"), Box::new(move |k| {
                    let mut root = Element::new("w:comments");
                    for n in list {
                        let mut kids = Vec::new();
                        content::write_blocks(&n.content, k, &mut kids);
                        root.children.push(Node::Element(Element { name: "w:comment".into(), attrs: n.attrs.clone(), children: kids }));
                    }
                    root
                })));
            }
            for (obj, file, kind, ct) in [
                (&self.comments_extended, "commentsExtended.xml", "commentsExtended", format!("{CT_WML}commentsExtended+xml")),
                (&self.comments_ids, "commentsIds.xml", "commentsIds", format!("{CT_WML}commentsIds+xml")),
                (&self.comments_extensible, "commentsExtensible.xml", "commentsExtensible", format!("{CT_WML}commentsExtensible+xml")),
                (&self.settings, "settings.xml", "settings", format!("{CT_WML}settings+xml")),
                (&self.web_settings, "webSettings.xml", "webSettings", format!("{CT_WML}webSettings+xml")),
                (&self.font_table, "fontTable.xml", "fontTable", format!("{CT_WML}fontTable+xml")),
                (&self.people, "people.xml", "people", format!("{CT_WML}people+xml")),
            ] {
                if let Some(o) = obj {
                    typed.push((file, kind, ct, Box::new(move |k| content::write_object(o, k))));
                }
            }
            for (file, kind, ct, build) in typed {
                let name = out.unique_name(file);
                let mut sink = PartRels::new(&mut out, &name, Vec::new());
                let root = build(&mut sink);
                let rels = std::mem::take(&mut sink.rels);
                drop(sink);
                out.put(&name, &ct, xml::write(&root, true).into_bytes());
                if !rels.is_empty() {
                    out.files.insert(rels_path(&name), rels_doc(&rels));
                }
                let rel_type = match kind {
                    "stylesWithEffects" => REL_SWE.to_string(),
                    "people" => REL_PEOPLE.to_string(),
                    "commentsExtended" => "http://schemas.microsoft.com/office/2011/relationships/commentsExtended".to_string(),
                    "commentsIds" => "http://schemas.microsoft.com/office/2016/09/relationships/commentsIds".to_string(),
                    "commentsExtensible" => "http://schemas.microsoft.com/office/2018/08/relationships/commentsExtensible".to_string(),
                    k => format!("{REL_BASE}{k}"),
                };
                doc_part_names.push((rel_type, name));
            }
        }

        // The document itself.
        let mut sink = PartRels::new(&mut out, &main, taken);
        for (rel_type, name) in &doc_part_names {
            sink.link(rel_type, name);
        }
        for p in self.theme.iter().chain(self.custom_xml.iter()).chain(self.glossary.iter()).chain(self.app_data.iter()) {
            sink.add(p);
        }
        let mut body = Vec::new();
        content::write_blocks(&self.body, &mut sink, &mut body);
        // Section breaks inside paragraphs and the final section name their
        // headers and footers by index.
        let link_hf = |e: &mut Element, sink: &mut PartRels| {
            for c in e.elements_mut() {
                if c.name == "w:headerReference" || c.name == "w:footerReference" {
                    if let Some(i) = c.attr("r:id").and_then(|v| v.strip_prefix('#')).and_then(|n| n.parse::<usize>().ok()) {
                        if let Some((name, header)) = hf_names.get(i) {
                            let rel = if *header { REL_HEADER } else { REL_FOOTER };
                            let id = sink.link(rel, name);
                            c.set_attr("r:id", id);
                        }
                    }
                }
            }
        };
        for n in body.iter_mut() {
            if let Node::Element(e) = n {
                visit_sectprs(e, &mut |s| link_hf(s, &mut sink));
            }
        }
        if let Some(s) = &self.section {
            let mut se = s.element().clone();
            link_hf(&mut se, &mut sink);
            let at = self.section_at.unwrap_or(body.len()).min(body.len());
            body.insert(at, Node::Element(se));
        }
        for (name, rel) in &note_parts {
            sink.link(rel, name);
        }
        let mut rels = std::mem::take(&mut sink.rels);
        drop(sink);
        doc_rels.append(&mut rels);
        let mut children = Vec::new();
        if let Some(bg) = &self.background {
            children.push(Node::Element(bg.clone()));
        }
        children.push(Node::Element(root_with("w:body", body)));
        let mut root = root_with("w:document", children);
        // Declare the extension namespaces Word may meet so mc:Ignorable holds.
        root.attrs.push(("mc:Ignorable".into(), "w14 w15 wp14 w16se w16cid w16 w16cex w16sdtdh w16du".into()));
        for p in ["w14", "w15", "wp14", "w16se", "w16cid", "w16", "w16cex", "w16sdtdh", "w16du"] {
            root.attrs.push((format!("xmlns:{p}"), xml::uri_for(p).unwrap().to_string()));
        }
        out.put(&main, CT_DOCUMENT, xml::write(&root, true).into_bytes());
        out.files.insert(rels_path(&main), rels_doc(&doc_rels));

        // [Content_Types].xml
        let mut defaults: BTreeMap<String, String> = BTreeMap::new();
        defaults.insert("rels".into(), CT_RELS.into());
        defaults.insert("xml".into(), "application/xml".into());
        let mut overrides = Vec::new();
        for (name, ct) in &out.types {
            let ext = name.rsplit('.').next().unwrap_or("").to_ascii_lowercase();
            if ext != "xml" && !ext.is_empty() && !name.ends_with(".rels") {
                match defaults.get(&ext) {
                    None => {
                        defaults.insert(ext, ct.clone());
                    }
                    Some(d) if d == ct => {}
                    Some(_) => overrides.push((name.clone(), ct.clone())),
                }
            } else if !ct.is_empty() {
                overrides.push((name.clone(), ct.clone()));
            }
        }
        let mut ct = String::from("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">");
        for (ext, t) in &defaults {
            ct.push_str(&format!(r#"<Default Extension="{}" ContentType="{}"/>"#, xml::escape(ext, true), xml::escape(t, true)));
        }
        for (name, t) in &overrides {
            ct.push_str(&format!(r#"<Override PartName="/{}" ContentType="{}"/>"#, xml::escape(name, true), xml::escape(t, true)));
        }
        ct.push_str("</Types>");
        out.files.insert("[Content_Types].xml".into(), ct.into_bytes());
        Ok(Package { files: out.files }.write())
    }
}

/// Visit every sectPr in a written subtree (paragraph section breaks).
fn visit_sectprs(e: &mut Element, f: &mut dyn FnMut(&mut Element)) {
    if e.name == "w:sectPr" {
        f(e);
        return;
    }
    for c in e.elements_mut() {
        visit_sectprs(c, f);
    }
}
