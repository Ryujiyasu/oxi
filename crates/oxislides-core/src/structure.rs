// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! Structural edits on a .pptx package: whole slides (duplicate, delete,
//! move) and whole shapes (insert a text box, delete, fill colour).
//!
//! Every function takes the package bytes and returns new package bytes, the
//! same contract as `PptxEditor::save`, so an editor can apply pending text
//! edits first and then a structural edit, and re-read the result.
//!
//! Shape indices count `p:sp` and `p:pic` elements inside `p:spTree` in
//! document order -- the same count `PptxEditor` uses for its text, format and
//! geometry edits, so an index a page got from one stays valid for the other.

use std::io::{Cursor, Read, Write};

use zip::write::SimpleFileOptions;
use zip::{CompressionMethod, ZipArchive, ZipWriter};

use crate::parser::PptxError;

const REL_SLIDE: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide";
const CT_SLIDE: &str = "application/vnd.openxmlformats-officedocument.presentationml.slide+xml";

/// One part of the package: either still the source archive's entry `orig`
/// (copied through compressed, never re-encoded) or new bytes.
struct Part {
    name: String,
    orig: Option<usize>,
    data: Option<Vec<u8>>,
}

/// The package as an ordered list of parts. Untouched parts -- pictures,
/// media, fonts, most of the XML -- are copied raw on write, so a structural
/// edit costs the parts it changes, not the size of the deck.
struct Package {
    src: Vec<u8>,
    parts: Vec<Part>,
}

fn zerr(e: impl std::fmt::Display) -> PptxError {
    PptxError::InvalidData(e.to_string())
}

impl Package {
    fn read(data: &[u8]) -> Result<Self, PptxError> {
        let zip = ZipArchive::new(Cursor::new(data)).map_err(zerr)?;
        let parts = (0..zip.len())
            .filter_map(|i| zip.name_for_index(i).map(|n| Part { name: n.to_string(), orig: Some(i), data: None }))
            .collect();
        Ok(Package { src: data.to_vec(), parts })
    }

    fn write(&self) -> Result<Vec<u8>, PptxError> {
        let mut zip = ZipArchive::new(Cursor::new(self.src.as_slice())).map_err(zerr)?;
        let mut out = Vec::new();
        {
            let mut w = ZipWriter::new(Cursor::new(&mut out));
            for p in &self.parts {
                match (&p.data, p.orig) {
                    (Some(data), orig) => {
                        let stored = orig
                            .and_then(|i| zip.by_index_raw(i).ok().map(|f| f.compression() == CompressionMethod::Stored))
                            .unwrap_or(false);
                        let comp = if stored { CompressionMethod::Stored } else { CompressionMethod::Deflated };
                        w.start_file(p.name.as_str(), SimpleFileOptions::default().compression_method(comp)).map_err(zerr)?;
                        w.write_all(data).map_err(zerr)?;
                    }
                    (None, Some(i)) => {
                        let f = zip.by_index_raw(i).map_err(zerr)?;
                        w.raw_copy_file(f).map_err(zerr)?;
                    }
                    (None, None) => {}
                }
            }
            w.finish().map_err(zerr)?;
        }
        Ok(out)
    }

    fn get(&self, name: &str) -> Option<Vec<u8>> {
        let p = self.parts.iter().find(|p| p.name == name)?;
        if let Some(d) = &p.data {
            return Some(d.clone());
        }
        let mut zip = ZipArchive::new(Cursor::new(self.src.as_slice())).ok()?;
        let mut f = zip.by_index(p.orig?).ok()?;
        let mut buf = Vec::new();
        f.read_to_end(&mut buf).ok()?;
        Some(buf)
    }

    fn text(&self, name: &str) -> Result<String, PptxError> {
        let b = self.get(name).ok_or_else(|| PptxError::InvalidData(format!("missing part {name}")))?;
        Ok(String::from_utf8_lossy(&b).into_owned())
    }

    fn set(&mut self, name: &str, data: Vec<u8>) {
        if let Some(p) = self.parts.iter_mut().find(|p| p.name == name) {
            p.data = Some(data);
        } else {
            self.parts.push(Part { name: name.to_string(), orig: None, data: Some(data) });
        }
    }

    fn remove(&mut self, name: &str) {
        self.parts.retain(|p| p.name != name);
    }

    fn has(&self, name: &str) -> bool {
        self.parts.iter().any(|p| p.name == name)
    }
}

/// One `Relationship` element of a .rels part.
#[derive(Clone, Debug)]
struct Rel {
    id: String,
    rtype: String,
    target: String,
    /// The whole element text, so unknown attributes (TargetMode) survive.
    raw: String,
}

fn attr(tag: &str, name: &str) -> Option<String> {
    let pat = format!(" {name}=\"");
    let i = tag.find(&pat)? + pat.len();
    let j = tag[i..].find('"')? + i;
    Some(tag[i..j].to_string())
}

fn parse_rels(xml: &str) -> Vec<Rel> {
    let mut out = Vec::new();
    let mut rest = xml;
    while let Some(i) = rest.find("<Relationship ") {
        let tail = &rest[i..];
        let end = tail.find("/>").map(|e| e + 2).or_else(|| tail.find("</Relationship>").map(|e| e + 15));
        let Some(end) = end else { break };
        let raw = &tail[..end];
        out.push(Rel {
            id: attr(raw, "Id").unwrap_or_default(),
            rtype: attr(raw, "Type").unwrap_or_default(),
            target: attr(raw, "Target").unwrap_or_default(),
            raw: raw.to_string(),
        });
        rest = &tail[end..];
    }
    out
}

fn write_rels(original: &str, rels: &[Rel]) -> String {
    let head_end = original.find("<Relationships").and_then(|i| original[i..].find('>').map(|j| i + j + 1));
    let head = match head_end {
        Some(h) => original[..h].to_string(),
        None => "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">".to_string(),
    };
    let mut s = head;
    for r in rels {
        s.push_str(&r.raw);
    }
    s.push_str("</Relationships>");
    s
}

/// The .rels part that belongs to a part ("ppt/slides/slide3.xml" ->
/// "ppt/slides/_rels/slide3.xml.rels").
fn rels_of(part: &str) -> String {
    match part.rfind('/') {
        Some(i) => format!("{}/_rels/{}.rels", &part[..i], &part[i + 1..]),
        None => format!("_rels/{part}.rels"),
    }
}

/// Resolve a relationship target against the directory of its source part.
fn resolve(source_part: &str, target: &str) -> String {
    if let Some(t) = target.strip_prefix('/') {
        return t.to_string();
    }
    let mut dir: Vec<&str> = match source_part.rfind('/') {
        Some(i) => source_part[..i].split('/').collect(),
        None => Vec::new(),
    };
    for seg in target.split('/') {
        match seg {
            ".." => {
                dir.pop();
            }
            "." | "" => {}
            s => dir.push(s),
        }
    }
    dir.join("/")
}

/// A target for `to_part` written relative to the directory of `from_part`.
fn relative(from_part: &str, to_part: &str) -> String {
    let from_dir: Vec<&str> = match from_part.rfind('/') {
        Some(i) => from_part[..i].split('/').collect(),
        None => Vec::new(),
    };
    let to: Vec<&str> = to_part.split('/').collect();
    let mut common = 0;
    while common < from_dir.len() && common + 1 < to.len() && from_dir[common] == to[common] {
        common += 1;
    }
    let mut out: Vec<String> = Vec::new();
    for _ in common..from_dir.len() {
        out.push("..".to_string());
    }
    for seg in &to[common..] {
        out.push((*seg).to_string());
    }
    out.join("/")
}

fn next_rid(rels: &[Rel]) -> String {
    let max = rels
        .iter()
        .filter_map(|r| r.id.strip_prefix("rId").and_then(|n| n.parse::<u32>().ok()))
        .max()
        .unwrap_or(0);
    format!("rId{}", max + 1)
}

/// A fresh part name "dir/stemN.ext" not yet in the package.
fn fresh_name(pkg: &Package, dir: &str, stem: &str, ext: &str) -> String {
    let mut n = 1;
    loop {
        let name = format!("{dir}/{stem}{n}.{ext}");
        if !pkg.has(&name) {
            return name;
        }
        n += 1;
    }
}

fn add_override(ct: &str, part: &str, content_type: &str) -> String {
    let entry = format!("<Override PartName=\"/{part}\" ContentType=\"{content_type}\"/>");
    match ct.rfind("</Types>") {
        Some(i) => format!("{}{}{}", &ct[..i], entry, &ct[i..]),
        None => ct.to_string(),
    }
}

fn remove_override(ct: &str, part: &str) -> String {
    let key = format!("PartName=\"/{part}\"");
    let Some(i) = ct.find(&key) else { return ct.to_string() };
    let start = ct[..i].rfind("<Override").unwrap_or(i);
    let end = ct[i..].find("/>").map(|e| i + e + 2).unwrap_or(i);
    format!("{}{}", &ct[..start], &ct[end..])
}

/// The `p:sldId` entries of presentation.xml as (full element text, r:id), in order,
/// with the byte span of the whole `p:sldIdLst` content.
type SlideIdList = (usize, usize, Vec<(String, String)>);

fn slide_ids(pres: &str) -> Option<SlideIdList> {
    let lst = pres.find("<p:sldIdLst")?;
    let open_end = lst + pres[lst..].find('>')? + 1;
    let close = open_end + pres[open_end..].find("</p:sldIdLst>")?;
    let mut items = Vec::new();
    let mut rest = &pres[open_end..close];
    while let Some(i) = rest.find("<p:sldId ") {
        let tail = &rest[i..];
        let end = tail.find("/>").map(|e| e + 2)?;
        let raw = tail[..end].to_string();
        let rid = attr(&raw, "r:id").unwrap_or_default();
        items.push((raw, rid));
        rest = &tail[end..];
    }
    Some((open_end, close, items))
}

/// The slide part names in presentation order.
pub fn slide_parts(data: &[u8]) -> Result<Vec<String>, PptxError> {
    let pkg = Package::read(data)?;
    slide_parts_of(&pkg)
}

fn slide_parts_of(pkg: &Package) -> Result<Vec<String>, PptxError> {
    let pres = pkg.text("ppt/presentation.xml")?;
    let rels = parse_rels(&pkg.text("ppt/_rels/presentation.xml.rels")?);
    let (_, _, items) = slide_ids(&pres).ok_or_else(|| PptxError::InvalidData("no sldIdLst".into()))?;
    Ok(items
        .iter()
        .filter_map(|(_, rid)| rels.iter().find(|r| &r.id == rid))
        .map(|r| resolve("ppt/presentation.xml", &r.target))
        .collect())
}

fn bad(msg: &str) -> PptxError {
    PptxError::InvalidData(msg.to_string())
}

/// Remove slide `index` (0-based). Its part, its relationships and its content
/// type entry go; parts only it pointed to (notes, charts) are left orphaned,
/// which PowerPoint ignores.
pub fn delete_slide(data: &[u8], index: usize) -> Result<Vec<u8>, PptxError> {
    let mut pkg = Package::read(data)?;
    let pres = pkg.text("ppt/presentation.xml")?;
    let (open, close, mut items) = slide_ids(&pres).ok_or_else(|| bad("no sldIdLst"))?;
    if index >= items.len() {
        return Err(bad("slide index out of range"));
    }
    if items.len() == 1 {
        return Err(bad("cannot delete the only slide"));
    }
    let (_, rid) = items.remove(index);
    let new_pres = format!("{}{}{}", &pres[..open], items.iter().map(|i| i.0.as_str()).collect::<String>(), &pres[close..]);
    pkg.set("ppt/presentation.xml", new_pres.into_bytes());

    let rels_xml = pkg.text("ppt/_rels/presentation.xml.rels")?;
    let mut rels = parse_rels(&rels_xml);
    let target = rels.iter().find(|r| r.id == rid).map(|r| resolve("ppt/presentation.xml", &r.target));
    rels.retain(|r| r.id != rid);
    pkg.set("ppt/_rels/presentation.xml.rels", write_rels(&rels_xml, &rels).into_bytes());

    if let Some(part) = target {
        pkg.remove(&part);
        pkg.remove(&rels_of(&part));
        let ct = pkg.text("[Content_Types].xml")?;
        pkg.set("[Content_Types].xml", remove_override(&ct, &part).into_bytes());
    }
    pkg.write()
}

/// Move slide `from` to position `to` (both 0-based, `to` counted after removal).
pub fn move_slide(data: &[u8], from: usize, to: usize) -> Result<Vec<u8>, PptxError> {
    let mut pkg = Package::read(data)?;
    let pres = pkg.text("ppt/presentation.xml")?;
    let (open, close, mut items) = slide_ids(&pres).ok_or_else(|| bad("no sldIdLst"))?;
    if from >= items.len() || to >= items.len() {
        return Err(bad("slide index out of range"));
    }
    let it = items.remove(from);
    items.insert(to, it);
    let new_pres = format!("{}{}{}", &pres[..open], items.iter().map(|i| i.0.as_str()).collect::<String>(), &pres[close..]);
    pkg.set("ppt/presentation.xml", new_pres.into_bytes());
    pkg.write()
}

/// The content type an Override gives `part`, if any.
fn override_type(ct: &str, part: &str) -> Option<String> {
    let key = format!("PartName=\"/{part}\"");
    let i = ct.find(&key)?;
    let start = ct[..i].rfind("<Override")?;
    let end = ct[i..].find("/>")? + i + 2;
    attr(&ct[start..end], "ContentType")
}

/// Relationship types whose target belongs to its source alone, so a copied
/// source needs its own copy: a chart and what hangs off it (embedded
/// workbook, style, colours, user shapes), SmartArt's data and drawing parts,
/// embedded OLE packages, and a notes page. Everything else (layout, pictures,
/// media, tags, hyperlinks) may be shared and is.
fn owned_rel(rtype: &str) -> bool {
    const OWNED: [&str; 11] = [
        "/chart", "/package", "/oleObject", "/chartUserShapes", "/diagramData", "/diagramDrawing",
        "/diagramLayout", "/diagramQuickStyle", "/diagramColors", "/notesSlide", "/chartStyle",
    ];
    OWNED.iter().any(|t| rtype.ends_with(t)) || rtype.ends_with("/chartColorStyle")
}

/// "ppt/charts/chart3.xml" -> ("ppt/charts", "chart", "xml").
fn split_part(part: &str) -> (String, String, String) {
    let (dir, file) = match part.rfind('/') {
        Some(i) => (&part[..i], &part[i + 1..]),
        None => ("", part),
    };
    let (base, ext) = match file.rfind('.') {
        Some(i) => (&file[..i], &file[i + 1..]),
        None => (file, ""),
    };
    let stem = base.trim_end_matches(|c: char| c.is_ascii_digit());
    (dir.to_string(), stem.to_string(), ext.to_string())
}

/// Copy `part` (and, recursively, the parts it owns) under fresh names; return
/// the copy's name. A relationship of type `.../slide` inside the copy (a notes
/// page pointing back at its slide) is re-aimed from `old_slide` to `new_slide`.
fn clone_part(pkg: &mut Package, ct: &mut String, part: &str, old_slide: &str, new_slide: &str) -> Option<String> {
    let data = pkg.get(part)?;
    let (dir, stem, ext) = split_part(part);
    let copy = fresh_name(pkg, &dir, &stem, &ext);
    pkg.set(&copy, data);
    if let Some(t) = override_type(ct, part) {
        *ct = add_override(ct, &copy, &t);
    }
    if let Some(rb) = pkg.get(&rels_of(part)) {
        let rels_xml = String::from_utf8_lossy(&rb).into_owned();
        let mut rels = parse_rels(&rels_xml);
        relink(pkg, ct, &mut rels, part, &copy, old_slide, new_slide);
        pkg.set(&rels_of(&copy), write_rels(&rels_xml, &rels).into_bytes());
    }
    Some(copy)
}

/// Re-aim the relationships of a copied part (`from` copied to `to`): owned
/// targets are cloned, a back-pointer to the old slide goes to the new one,
/// and shared targets keep pointing where they did.
fn relink(pkg: &mut Package, ct: &mut String, rels: &mut [Rel], from: &str, to: &str, old_slide: &str, new_slide: &str) {
    for r in rels.iter_mut() {
        if r.raw.contains("TargetMode=\"External\"") {
            continue;
        }
        let target = resolve(from, &r.target);
        let new_target = if r.rtype.ends_with("/slide") && target == old_slide {
            Some(new_slide.to_string())
        } else if owned_rel(&r.rtype) {
            clone_part(pkg, ct, &target, old_slide, new_slide)
        } else {
            None
        };
        if let Some(nt) = new_target {
            let rel_t = relative(to, &nt);
            r.raw = r.raw.replace(&format!("Target=\"{}\"", r.target), &format!("Target=\"{rel_t}\""));
            r.target = rel_t;
        }
    }
}

/// Copy slide `index` and insert the copy right after it, as PowerPoint's
/// Duplicate Slide does.
///
/// The copy shares its layout and pictures with the original. Parts that
/// belong to one slide alone -- its notes page, charts with their embedded
/// workbooks, SmartArt data -- are copied: two slides naming one of those make
/// PowerPoint refuse or repair the file.
pub fn duplicate_slide(data: &[u8], index: usize) -> Result<Vec<u8>, PptxError> {
    let mut pkg = Package::read(data)?;
    let parts = slide_parts_of(&pkg)?;
    let src = parts.get(index).cloned().ok_or_else(|| bad("slide index out of range"))?;
    let dst = fresh_name(&pkg, "ppt/slides", "slide", "xml");

    let slide_xml = pkg.get(&src).ok_or_else(|| bad("missing slide part"))?;
    pkg.set(&dst, slide_xml);
    let mut ct = pkg.text("[Content_Types].xml")?;
    ct = add_override(&ct, &dst, CT_SLIDE);

    if let Some(rels_bytes) = pkg.get(&rels_of(&src)) {
        let rels_xml = String::from_utf8_lossy(&rels_bytes).into_owned();
        let mut rels = parse_rels(&rels_xml);
        // Comments are about the original slide; the copy starts without them.
        rels.retain(|r| !r.rtype.ends_with("/comments"));
        relink(&mut pkg, &mut ct, &mut rels, &src, &dst, &src, &dst);
        pkg.set(&rels_of(&dst), write_rels(&rels_xml, &rels).into_bytes());
    }
    pkg.set("[Content_Types].xml", ct.into_bytes());

    let prels_xml = pkg.text("ppt/_rels/presentation.xml.rels")?;
    let mut prels = parse_rels(&prels_xml);
    let rid = next_rid(&prels);
    let target = relative("ppt/presentation.xml", &dst);
    prels.push(Rel {
        id: rid.clone(),
        rtype: REL_SLIDE.to_string(),
        target: target.clone(),
        raw: format!("<Relationship Id=\"{rid}\" Type=\"{REL_SLIDE}\" Target=\"{target}\"/>"),
    });
    pkg.set("ppt/_rels/presentation.xml.rels", write_rels(&prels_xml, &prels).into_bytes());

    let pres = pkg.text("ppt/presentation.xml")?;
    let (open, close, mut items) = slide_ids(&pres).ok_or_else(|| bad("no sldIdLst"))?;
    let next_id = items
        .iter()
        .filter_map(|(raw, _)| attr(raw, "id").and_then(|v| v.parse::<u32>().ok()))
        .max()
        .unwrap_or(255)
        .max(255)
        + 1;
    items.insert(index + 1, (format!("<p:sldId id=\"{next_id}\" r:id=\"{rid}\"/>"), rid));
    let new_pres = format!("{}{}{}", &pres[..open], items.iter().map(|i| i.0.as_str()).collect::<String>(), &pres[close..]);
    pkg.set("ppt/presentation.xml", new_pres.into_bytes());
    pkg.write()
}

/// Byte spans of the top-level shape elements a shape index addresses: every
/// `p:sp` / `p:pic` inside `p:spTree`, in document order (nested ones in
/// groups included, the same count `PptxEditor` keeps).
fn shape_spans(xml: &str) -> Vec<(usize, usize)> {
    let mut out = Vec::new();
    let Some(tree) = xml.find("<p:spTree") else { return out };
    let tree_end = xml[tree..].find("</p:spTree>").map(|e| tree + e).unwrap_or(xml.len());
    let mut pos = tree;
    while pos < tree_end {
        let next = ["<p:sp>", "<p:sp ", "<p:pic>", "<p:pic "]
            .iter()
            .filter_map(|t| xml[pos..tree_end].find(t).map(|i| (pos + i, *t)))
            .min_by_key(|x| x.0);
        let Some((start, tag)) = next else { break };
        let name = if tag.starts_with("<p:sp") { "p:sp" } else { "p:pic" };
        let close = format!("</{name}>");
        // Same-named elements do not nest (a p:sp holds no p:sp), so the first
        // close tag ends this one.
        let end = match xml[start..tree_end].find(&close) {
            Some(e) => start + e + close.len(),
            None => break,
        };
        out.push((start, end));
        pos = start + 1;
        // Skip past this element's opening tag so its children are scanned too,
        // exactly as the editor counts nested shapes.
        pos = xml[pos..].find('>').map(|g| pos + g + 1).unwrap_or(end);
    }
    out
}

fn slide_part(pkg: &Package, slide: usize) -> Result<String, PptxError> {
    slide_parts_of(pkg)?.get(slide).cloned().ok_or_else(|| bad("slide index out of range"))
}

/// Remove shape `shape` (a `p:sp` or `p:pic`) from slide `slide`.
pub fn delete_shape(data: &[u8], slide: usize, shape: usize) -> Result<Vec<u8>, PptxError> {
    let mut pkg = Package::read(data)?;
    let part = slide_part(&pkg, slide)?;
    let xml = pkg.text(&part)?;
    let spans = shape_spans(&xml);
    let (s, e) = *spans.get(shape).ok_or_else(|| bad("shape index out of range"))?;
    let new_xml = format!("{}{}", &xml[..s], &xml[e..]);
    pkg.set(&part, new_xml.into_bytes());
    pkg.write()
}

fn xml_escape(s: &str) -> String {
    s.replace('&', "&amp;").replace('<', "&lt;").replace('>', "&gt;").replace('"', "&quot;")
}

fn max_shape_id(xml: &str) -> u32 {
    let mut max = 1;
    let mut rest = xml;
    while let Some(i) = rest.find("<p:cNvPr ") {
        let tail = &rest[i..];
        let end = tail.find('>').unwrap_or(tail.len());
        if let Some(v) = attr(&tail[..end], "id").and_then(|v| v.parse::<u32>().ok()) {
            max = max.max(v);
        }
        rest = &tail[end.min(tail.len())..];
        if end == 0 {
            break;
        }
    }
    max
}

/// Append a text box to slide `slide`. Geometry is in POINTS; the text is split
/// into paragraphs on '\n'. Returns the new package and the new shape's index.
#[allow(clippy::too_many_arguments)]
pub fn add_text_box(
    data: &[u8],
    slide: usize,
    x: f32,
    y: f32,
    w: f32,
    h: f32,
    text: &str,
    font_size: Option<f32>,
) -> Result<(Vec<u8>, usize), PptxError> {
    let mut pkg = Package::read(data)?;
    let part = slide_part(&pkg, slide)?;
    let xml = pkg.text(&part)?;
    let index = shape_spans(&xml).len();
    let id = max_shape_id(&xml) + 1;
    let emu = |v: f32| (v as f64 * 12700.0).round() as i64;
    let rpr = match font_size {
        Some(sz) => format!("<a:rPr lang=\"en-US\" sz=\"{}\" dirty=\"0\"/>", (sz * 100.0).round() as i64),
        None => "<a:rPr lang=\"en-US\" dirty=\"0\"/>".to_string(),
    };
    let paras: String = text
        .split('\n')
        .map(|line| {
            if line.is_empty() {
                "<a:p><a:endParaRPr lang=\"en-US\" dirty=\"0\"/></a:p>".to_string()
            } else {
                format!("<a:p><a:r>{rpr}<a:t>{}</a:t></a:r></a:p>", xml_escape(line))
            }
        })
        .collect();
    let sp = format!(
        "<p:sp><p:nvSpPr><p:cNvPr id=\"{id}\" name=\"TextBox {id}\"/><p:cNvSpPr txBox=\"1\"/><p:nvPr/></p:nvSpPr>\
<p:spPr><a:xfrm><a:off x=\"{}\" y=\"{}\"/><a:ext cx=\"{}\" cy=\"{}\"/></a:xfrm><a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom><a:noFill/></p:spPr>\
<p:txBody><a:bodyPr wrap=\"square\" rtlCol=\"0\"><a:spAutoFit/></a:bodyPr><a:lstStyle/>{paras}</p:txBody></p:sp>",
        emu(x),
        emu(y),
        emu(w.max(1.0)),
        emu(h.max(1.0))
    );
    let end = xml.find("</p:spTree>").ok_or_else(|| bad("slide has no spTree"))?;
    let new_xml = format!("{}{}{}", &xml[..end], sp, &xml[end..]);
    pkg.set(&part, new_xml.into_bytes());
    Ok((pkg.write()?, index))
}

/// Set (Some "RRGGBB") or clear to no fill (None) the fill of shape `shape`.
/// The new fill replaces any existing fill child of `p:spPr` and sits where
/// the schema wants it: after `a:xfrm` and the geometry.
pub fn set_shape_fill(data: &[u8], slide: usize, shape: usize, color: Option<&str>) -> Result<Vec<u8>, PptxError> {
    let mut pkg = Package::read(data)?;
    let part = slide_part(&pkg, slide)?;
    let xml = pkg.text(&part)?;
    let (s, e) = *shape_spans(&xml).get(shape).ok_or_else(|| bad("shape index out of range"))?;
    let shape_xml = &xml[s..e];
    let sp_open = shape_xml.find("<p:spPr").ok_or_else(|| bad("shape has no spPr"))?;
    let new_fill = match color {
        Some(c) => format!("<a:solidFill><a:srgbClr val=\"{}\"/></a:solidFill>", c.trim_start_matches('#').to_uppercase()),
        None => "<a:noFill/>".to_string(),
    };
    let patched = if shape_xml[sp_open..].starts_with("<p:spPr/>") {
        format!("{}<p:spPr>{new_fill}</p:spPr>{}", &shape_xml[..sp_open], &shape_xml[sp_open + 9..])
    } else {
        let body_start = sp_open + shape_xml[sp_open..].find('>').ok_or_else(|| bad("bad spPr"))? + 1;
        let body_end = body_start + shape_xml[body_start..].find("</p:spPr>").ok_or_else(|| bad("bad spPr"))?;
        let body = &shape_xml[body_start..body_end];
        // Drop existing fill children (top level of spPr only).
        let mut kept = String::new();
        let mut rest = body;
        let mut insert_at: Option<usize> = None;
        let fills = ["a:noFill", "a:solidFill", "a:gradFill", "a:blipFill", "a:pattFill", "a:grpFill"];
        while !rest.is_empty() {
            let Some(lt) = rest.find('<') else {
                kept.push_str(rest);
                break;
            };
            kept.push_str(&rest[..lt]);
            let tail = &rest[lt..];
            let name_end = tail[1..].find([' ', '>', '/']).map(|i| i + 1).unwrap_or(tail.len());
            let name = &tail[1..name_end];
            // Extent of this child element.
            let child_end = if tail[..tail.find('>').unwrap_or(0) + 1].ends_with("/>") {
                tail.find('>').unwrap() + 1
            } else {
                let close = format!("</{name}>");
                // Children of spPr may nest same-named elements only inside
                // a:extLst; count depth to be safe.
                let open = format!("<{name}");
                let mut depth = 0i32;
                let mut i = 0usize;
                let mut endpos = tail.len();
                while i < tail.len() {
                    if tail[i..].starts_with(&close) {
                        depth -= 1;
                        if depth == 0 {
                            endpos = i + close.len();
                            break;
                        }
                        i += close.len();
                    } else if tail[i..].starts_with(&open)
                        && tail[i + open.len()..].starts_with([' ', '>', '/'])
                    {
                        let gt = tail[i..].find('>').map(|g| i + g).unwrap_or(tail.len() - 1);
                        if !tail[..=gt].ends_with("/>") {
                            depth += 1;
                        }
                        i = gt + 1;
                    } else {
                        i += tail[i..].chars().next().map(|c| c.len_utf8()).unwrap_or(1);
                    }
                }
                endpos
            };
            if fills.contains(&name) {
                if insert_at.is_none() {
                    insert_at = Some(kept.len());
                }
            } else {
                kept.push_str(&tail[..child_end]);
                if name == "a:xfrm" || name == "a:custGeom" || name == "a:prstGeom" {
                    insert_at = Some(kept.len());
                }
            }
            rest = &tail[child_end..];
        }
        let at = insert_at.unwrap_or(0);
        let new_body = format!("{}{}{}", &kept[..at], new_fill, &kept[at..]);
        format!("{}{}{}", &shape_xml[..body_start], new_body, &shape_xml[body_end..])
    };
    let new_xml = format!("{}{}{}", &xml[..s], patched, &xml[e..]);
    pkg.set(&part, new_xml.into_bytes());
    pkg.write()
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn relative_and_resolve_round_trip() {
        assert_eq!(relative("ppt/presentation.xml", "ppt/slides/slide3.xml"), "slides/slide3.xml");
        assert_eq!(relative("ppt/slides/slide2.xml", "ppt/charts/chart4.xml"), "../charts/chart4.xml");
        assert_eq!(resolve("ppt/slides/slide2.xml", "../charts/chart4.xml"), "ppt/charts/chart4.xml");
        assert_eq!(rels_of("ppt/slides/slide3.xml"), "ppt/slides/_rels/slide3.xml.rels");
    }

    #[test]
    fn shape_spans_count_sp_and_pic_including_group_children() {
        let xml = "<p:spTree><p:sp><a/></p:sp><p:grpSp><p:sp>x</p:sp><p:pic>y</p:pic></p:grpSp><p:pic>z</p:pic></p:spTree>";
        let s = shape_spans(xml);
        assert_eq!(s.len(), 4);
        assert_eq!(&xml[s[1].0..s[1].1], "<p:sp>x</p:sp>");
        assert_eq!(&xml[s[3].0..s[3].1], "<p:pic>z</p:pic>");
    }

    #[test]
    fn content_type_override_add_and_remove() {
        let ct = "<Types><Default Extension=\"xml\" ContentType=\"x\"/></Types>";
        let added = add_override(ct, "ppt/slides/slide9.xml", CT_SLIDE);
        assert!(added.contains("PartName=\"/ppt/slides/slide9.xml\""));
        assert_eq!(remove_override(&added, "ppt/slides/slide9.xml"), ct);
    }

    fn texts(data: &[u8]) -> Vec<String> {
        let pres = crate::parser::parse_pptx(data).expect("parse");
        pres.slides
            .iter()
            .map(|s| {
                s.shapes
                    .iter()
                    .filter_map(|sh| match &sh.content {
                        crate::ir::ShapeContent::TextBox { paragraphs } | crate::ir::ShapeContent::AutoShape { paragraphs, .. } => Some(
                            paragraphs.iter().flat_map(|p| p.runs.iter().map(|r| r.text.clone())).collect::<String>(),
                        ),
                        _ => None,
                    })
                    .collect::<Vec<_>>()
                    .join("|")
            })
            .collect()
    }

    const DECK: &[u8] = include_bytes!("../../../tests/fixtures/multi_slide.pptx");

    #[test]
    fn duplicate_inserts_copy_after_source_without_notes() {
        let before = texts(DECK);
        let out = duplicate_slide(DECK, 0).unwrap();
        let after = texts(&out);
        assert_eq!(after.len(), before.len() + 1);
        assert_eq!(after[0], before[0]);
        assert_eq!(after[1], before[0]);
        assert_eq!(after[2..], before[1..]);
        let parts = slide_parts(&out).unwrap();
        assert_eq!(parts[1], "ppt/slides/slide4.xml");
        let pkg = Package::read(&out).unwrap();
        assert!(pkg.text("[Content_Types].xml").unwrap().contains("/ppt/slides/slide4.xml"));
        let pres = pkg.text("ppt/presentation.xml").unwrap();
        let (_, _, ids) = slide_ids(&pres).unwrap();
        let mut nums: Vec<u32> = ids.iter().map(|(r, _)| attr(r, "id").unwrap().parse().unwrap()).collect();
        nums.sort();
        nums.dedup();
        assert_eq!(nums.len(), ids.len());
    }

    #[test]
    fn delete_and_move_reorder_slides() {
        let before = texts(DECK);
        let del = texts(&delete_slide(DECK, 1).unwrap());
        assert_eq!(del, vec![before[0].clone(), before[2].clone()]);
        let mv = texts(&move_slide(DECK, 2, 0).unwrap());
        assert_eq!(mv, vec![before[2].clone(), before[0].clone(), before[1].clone()]);
        assert!(delete_slide(DECK, 9).is_err());
    }

    #[test]
    fn add_text_box_then_fill_then_delete_shape() {
        let n0 = crate::parser::parse_pptx(DECK).unwrap().slides[0].shapes.len();
        let (out, idx) = add_text_box(DECK, 0, 100.0, 120.0, 200.0, 40.0, "Hello & <you>
line 2", Some(24.0)).unwrap();
        let pres = crate::parser::parse_pptx(&out).unwrap();
        assert_eq!(pres.slides[0].shapes.len(), n0 + 1);
        let t = texts(&out); assert!(t[0].contains("Hello & <you>line 2"), "{t:?}");
        let filled = set_shape_fill(&out, 0, idx, Some("#ff8800")).unwrap();
        let xml = Package::read(&filled).unwrap().text("ppt/slides/slide1.xml").unwrap();
        assert!(xml.contains("</a:prstGeom><a:solidFill><a:srgbClr val=\"FF8800\"/></a:solidFill></p:spPr>"));
        assert!(!xml.contains("<a:noFill/></p:spPr>"));
        let gone = delete_shape(&filled, 0, idx).unwrap();
        assert_eq!(texts(&gone), texts(DECK));
    }

    #[test]
    fn duplicate_copies_chart_and_its_workbook() {
        let deck: &[u8] = include_bytes!("../../../tests/fixtures/pptx_chart_surface.pptx");
        let out = duplicate_slide(deck, 0).unwrap();
        let pkg = Package::read(&out).unwrap();
        let rels = parse_rels(&pkg.text("ppt/slides/_rels/slide7.xml.rels").unwrap());
        let chart = rels.iter().find(|r| r.rtype.ends_with("/chart")).map(|r| resolve("ppt/slides/slide7.xml", &r.target)).unwrap();
        assert_eq!(chart, "ppt/charts/chart7.xml");
        let crels = parse_rels(&pkg.text("ppt/charts/_rels/chart7.xml.rels").unwrap());
        let wb = resolve("ppt/charts/chart7.xml", &crels[0].target);
        assert_eq!(wb, "ppt/embeddings/Microsoft_Excel_Sheet7.xlsx");
        assert!(pkg.has(&wb));
        assert!(pkg.text("[Content_Types].xml").unwrap().contains("/ppt/charts/chart7.xml"));
        assert_eq!(crate::parser::parse_pptx(&out).unwrap().slides.len(), 7);
    }
}
