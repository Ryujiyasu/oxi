// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! Export a .docx laid out by Oxi as a [DoclingDocument] (schema 1.10.0) whose
//! items carry page and bounding-box provenance.
//!
//! Docling reads a .docx as a tree of text and leaves every item's `prov`
//! empty, because nothing lays the document out. Oxi does lay it out, so each
//! item here gets one `ProvenanceItem` per page it appears on.
//!
//! Conventions (fixed here so consumers can rely on them):
//! - Coordinates are PDF points, origin at the page's top-left
//!   (`coord_origin: "TOPLEFT"`); `pages[n].size` is in the same unit.
//! - A text box is the union of the item's line boxes on that page; a list
//!   item's box includes its number or bullet.
//! - `charspan` indexes the item's `orig` in Unicode code points (what Python
//!   `str` slicing uses). `orig` is the concatenation of the paragraph's runs as
//!   Word shows them in its final view: deleted revisions and hidden text are
//!   left out. `text` equals `orig`.
//! - A table's and a picture's `charspan` is `[0, 0]`.
//!
//! Structure and text follow docling's own docx output: section headers with
//! levels (numbered ones carry their number, "1.2 Title"), list groups,
//! `page_header` / `page_footer` furniture (one item per page), an `inline`
//! group of text and `formula` items for a paragraph with equations, one
//! `formula` item per display equation, `textbox` groups after the paragraph
//! they anchor to, table cells joined with newlines and inline equations as
//! `$…$`. Formula text is produced by a port of docling's OMML converter from
//! the document XML, so it matches docling character for character.
//!
//! [DoclingDocument]: https://github.com/docling-project/docling-core

mod math;
mod omml;
pub mod annotate;

use std::collections::{BTreeMap, HashMap, VecDeque};
use std::io::Read;

use oxidocs_core::ir::{Block, Document, Paragraph, Run, ShowRevisions, StyleSheet, Table, TextBox};
use oxidocs_core::layout::{LayoutContent, LayoutEngine, LayoutResult};
use serde_json::{json, Map, Value};

pub const SCHEMA_VERSION: &str = "1.10.0";
const DOCX_MIME: &str = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";

/// What the export could and could not place.
#[derive(Debug, Default, Clone)]
pub struct Stats {
    pub pages: usize,
    pub items: usize,
    pub items_with_prov: usize,
    /// Painted text fragments that no exported item claimed.
    pub unplaced_fragments: usize,
}

/// Parse, lay out and export a .docx.
pub fn docx_to_docling(data: &[u8], filename: &str) -> Result<(Value, Stats), String> {
    let doc = oxidocs_core::parse_docx(data).map_err(|e| format!("parse: {e:?}"))?;
    let res = LayoutEngine::for_document(&doc)
        .with_show_revisions(ShowRevisions::Final)
        .layout(&doc);
    let tex = document_xml(data).map(|x| omml::collect(&x)).unwrap_or_default();
    Ok(export(&doc, &res, tex, filename, fnv1a(data)))
}

fn document_xml(data: &[u8]) -> Option<String> {
    let mut z = zip::ZipArchive::new(std::io::Cursor::new(data)).ok()?;
    let mut s = String::new();
    z.by_name("word/document.xml").ok()?.read_to_string(&mut s).ok()?;
    Some(s)
}

// ---------- geometry ----------

#[derive(Clone, Copy)]
pub(crate) struct Rect {
    pub(crate) l: f32,
    pub(crate) t: f32,
    pub(crate) r: f32,
    pub(crate) b: f32,
}

impl Rect {
    fn union(self, o: Rect) -> Rect {
        Rect { l: self.l.min(o.l), t: self.t.min(o.t), r: self.r.max(o.r), b: self.b.max(o.b) }
    }
    fn json(self) -> Value {
        let r2 = |v: f32| ((v as f64) * 100.0).round() / 100.0;
        json!({"l": r2(self.l), "t": r2(self.t), "r": r2(self.r), "b": r2(self.b), "coord_origin": "TOPLEFT"})
    }
}

pub(crate) fn prov(page: usize, rect: Rect, span: (usize, usize)) -> Value {
    json!({"page_no": page, "bbox": rect.json(), "charspan": [span.0, span.1]})
}

pub(crate) fn union_by_page(it: impl IntoIterator<Item = (usize, Rect)>) -> BTreeMap<usize, Rect> {
    let mut m: BTreeMap<usize, Rect> = BTreeMap::new();
    for (p, r) in it {
        m.entry(p).and_modify(|u| *u = u.union(r)).or_insert(r);
    }
    m
}

fn provs_of(it: impl IntoIterator<Item = (usize, Rect)>, span: (usize, usize)) -> Vec<Value> {
    union_by_page(it).into_iter().map(|(p, r)| prov(p, r, span)).collect()
}

// ---------- the layout, indexed ----------

pub(crate) type Key = (usize, usize); // (section, body block)

pub(crate) struct Frag {
    pub(crate) page: usize,
    pub(crate) rect: Rect,
    /// (run index, char offset in run, chars) when the fragment maps to source
    /// text; `None` for inline equations, which the layout paints per glyph.
    pub(crate) src: Option<(usize, usize, usize)>,
    /// Top-level table cell (row, cell index in row) for table fragments.
    pub(crate) cell: Option<(usize, usize)>,
    /// Paragraph index within that cell (flattened, document order).
    pub(crate) cpi: Option<usize>,
    /// Source characters this fragment paints (before case conversion).
    pub(crate) text: String,
}

/// A painted text fragment with no body block: display equations, text boxes,
/// footnotes. (List numbers and bullets are claimed by their item in `index`.)
pub(crate) struct Free {
    pub(crate) text: String,
    pub(crate) rect: Rect,
    /// Text boxes keep their run index; equation glyphs have none.
    pub(crate) has_run: bool,
    pub(crate) used: bool,
}

pub(crate) enum Ev {
    Body(Key),
    Free(usize),
}

#[derive(Default)]
pub(crate) struct Index {
    pub(crate) body: HashMap<Key, Vec<Frag>>,
    /// A list item's number or bullet, painted just before its first fragment.
    pub(crate) markers: HashMap<Key, Vec<(usize, Rect)>>,
    pub(crate) pictures: HashMap<Option<Key>, Vec<(usize, Rect)>>,
    pub(crate) furniture: BTreeMap<(usize, bool), Vec<(String, Rect)>>,
    /// Pictures painted in header/footer areas, in page order.
    pub(crate) furniture_pictures: Vec<(usize, Rect)>,
    pub(crate) free: Vec<Free>,
    pub(crate) events: BTreeMap<usize, Vec<Ev>>,
    pub(crate) page_size: Vec<(f32, f32)>,
}

pub(crate) fn index(doc: &Document, res: &LayoutResult) -> Index {
    let mut ix = Index::default();
    let mut last_body: Option<Key> = None;
    for (pi, pg) in res.pages.iter().enumerate() {
        let page = pi + 1;
        ix.page_size.push((pg.width, pg.height));
        let sec = pg.elements.iter().find_map(|e| e.source_container_index.filter(|_| e.paragraph_index.is_some()))
            .or(last_body.map(|k| k.0)).unwrap_or(0);
        let margin = doc.pages.get(sec).map(|s| (s.margin.top, s.margin.bottom)).unwrap_or((72.0, 72.0));
        let in_header = |r: &Rect| r.b <= margin.0 + 1.0;
        let in_footer = |r: &Rect| r.t >= pg.height - margin.1 - 1.0;
        let evs = ix.events.entry(page).or_default();
        // A short no-run fragment waiting to be a list marker.
        let mut pending: Option<usize> = None;
        for e in &pg.elements {
            match &e.content {
                LayoutContent::Text { text, .. } => {
                    let top = e.y - e.text_y_off;
                    let rect = Rect { l: e.x, t: top, r: e.x + e.width, b: top + e.height };
                    match (e.source_container_index, e.paragraph_index) {
                        (Some(c), Some(p)) => {
                            let key = (c, p);
                            let src = match (e.run_index, e.char_offset) {
                                (Some(r), Some(o)) => Some((r, o, e.source_char_len.unwrap_or_else(|| text.chars().count()))),
                                _ => None,
                            };
                            let cell = match e.cell_ancestor_path.first() {
                                Some(&(row, cell, _)) => Some((row, cell)),
                                None => e.cell_row_index.zip(e.cell_col_index),
                            };
                            let first = !ix.body.contains_key(&key);
                            if let (true, Some(m)) = (first, pending) {
                                let f = &mut ix.free[m];
                                if f.rect.t < rect.b && rect.t < f.rect.b && f.rect.l <= rect.l {
                                    f.used = true;
                                    ix.markers.entry(key).or_default().push((page, f.rect));
                                }
                            }
                            pending = None;
                            let cpi = if e.cell_ancestor_path.is_empty() { e.cell_paragraph_index } else { None };
                            let text = e.source_text.clone().unwrap_or_else(|| text.clone());
                            ix.body.entry(key).or_default().push(Frag { page, rect, src, cell, cpi, text });
                            evs.push(Ev::Body(key));
                            last_body = Some(key);
                        }
                        _ if text.trim().is_empty() => {}
                        _ if in_header(&rect) => ix.furniture.entry((page, true)).or_default().push((text.clone(), rect)),
                        _ if in_footer(&rect) => ix.furniture.entry((page, false)).or_default().push((text.clone(), rect)),
                        _ => {
                            let i = ix.free.len();
                            let has_run = e.run_index.is_some();
                            pending = (!has_run && text.chars().count() <= 12).then_some(i);
                            evs.push(Ev::Free(i));
                            ix.free.push(Free { text: text.clone(), rect, has_run, used: false });
                        }
                    }
                }
                LayoutContent::Image { .. } => {
                    let rect = Rect { l: e.x, t: e.y, r: e.x + e.width, b: e.y + e.height };
                    if !in_header(&rect) && !in_footer(&rect) {
                        ix.pictures.entry(last_body).or_default().push((page, rect));
                    } else {
                        ix.furniture_pictures.push((page, rect));
                    }
                }
                _ => {}
            }
        }
    }
    ix
}

pub(crate) fn norm(s: &str) -> String {
    s.chars().filter(|c| !c.is_whitespace()).collect()
}

/// Location of every character of `orig`, by aligning the painted text of
/// `frags` (paint order) to it from left to right. The layout's own run
/// indices can refer to runs it split internally (CJK, kinsoku), so the
/// characters themselves are the reliable link. Fragments without source text
/// (equation glyphs, markers) are skipped; so is text the paragraph does not
/// contain (field results, inserted hyphens).
pub(crate) fn align_locs(orig: &str, frags: &[&Frag]) -> Vec<Option<(usize, Rect)>> {
    let chars: Vec<char> = orig.chars().collect();
    let mut locs = vec![None; chars.len()];
    let mut ptr = 0usize;
    for f in frags {
        if f.src.is_none() {
            continue;
        }
        let t: Vec<char> = f.text.chars().map(|c| if matches!(c, '\x0C' | '\x0B' | '\r') { '\n' } else { c }).collect();
        if t.is_empty() {
            continue;
        }
        let window = (ptr + t.len() + 64).min(chars.len());
        let found = (ptr..window.saturating_sub(t.len() - 1)).find(|&q| chars[q..q + t.len()] == t[..]);
        let Some(q) = found else { continue };
        let w = f.rect.r - f.rect.l;
        let n = t.len() as f32;
        for k in 0..t.len() {
            let x0 = f.rect.l + w * k as f32 / n;
            let x1 = f.rect.l + w * (k + 1) as f32 / n;
            locs[q + k] = Some((f.page, Rect { l: x0, r: x1, ..f.rect }));
        }
        ptr = q + t.len();
    }
    locs
}

impl Index {
    /// Free no-run fragments painted between this block's neighbours, on
    /// every page: where a display equation's glyphs (and in-cell equations
    /// and numbers for a table) land, since they carry no block index.
    pub(crate) fn take_between(&mut self, key: Key) -> Vec<(usize, Rect)> {
        let mut out = vec![];
        for (&page, evs) in &self.events {
            let before = evs.iter().rposition(|e| matches!(e, Ev::Body(k) if *k < key));
            let after = evs.iter().position(|e| matches!(e, Ev::Body(k) if *k > key));
            if before.is_none() && after.is_none() && !evs.iter().any(|e| matches!(e, Ev::Body(k) if *k == key)) {
                continue;
            }
            let lo = before.map_or(0, |i| i + 1);
            let hi = after.unwrap_or(evs.len());
            for ev in &evs[lo..hi] {
                if let Ev::Free(i) = ev {
                    let f = &mut self.free[*i];
                    if !f.used && !f.has_run {
                        f.used = true;
                        out.push((page, f.rect));
                    }
                }
            }
        }
        out
    }

    /// Consecutive free fragments with a run index, on one page, whose text
    /// concatenates to `target`: a text box paragraph.
    pub(crate) fn take_text(&mut self, target: &str) -> Option<(usize, Rect)> {
        let want = norm(target);
        if want.is_empty() {
            return None;
        }
        for (&page, evs) in &self.events {
            let idx: Vec<usize> = evs.iter().filter_map(|e| match e {
                Ev::Free(i) if !self.free[*i].used && self.free[*i].has_run => Some(*i),
                _ => None,
            }).collect();
            for start in 0..idx.len() {
                let mut acc = String::new();
                let mut rect: Option<Rect> = None;
                for (n, &i) in idx[start..].iter().enumerate() {
                    let t = norm(&self.free[i].text);
                    if t.is_empty() {
                        continue;
                    }
                    acc.push_str(&t);
                    if !want.starts_with(&acc) {
                        break;
                    }
                    rect = Some(rect.map_or(self.free[i].rect, |r| r.union(self.free[i].rect)));
                    if acc == want {
                        for &j in &idx[start..=start + n] {
                            self.free[j].used = true;
                        }
                        return Some((page, rect.unwrap()));
                    }
                }
            }
        }
        None
    }

    /// First painted fragment of the nearest body block at or after `key`
    /// (then before it): the anchor of a floating object.
    pub(crate) fn anchor_rect(&self, key: Key, n_blocks: usize) -> Option<(usize, Rect)> {
        let after = (key.1..n_blocks).map(|b| (key.0, b));
        let before = (0..key.1).rev().map(|b| (key.0, b));
        after.chain(before).find_map(|k| self.body.get(&k).and_then(|v| v.first()).map(|f| (f.page, f.rect)))
    }
}

/// Split free fragments into `k` clusters at the largest vertical gaps: the
/// `k` equations of one display paragraph.
pub(crate) fn clusters(mut frags: Vec<(usize, Rect)>, k: usize) -> Vec<Vec<(usize, Rect)>> {
    if frags.is_empty() {
        return vec![];
    }
    frags.sort_by(|a, b| a.0.cmp(&b.0).then(a.1.t.total_cmp(&b.1.t)));
    let mut gaps: Vec<(f32, usize)> = (1..frags.len()).map(|i| {
        let gap = if frags[i].0 != frags[i - 1].0 { f32::MAX } else { frags[i].1.t - frags[i - 1].1.b };
        (gap, i)
    }).collect();
    gaps.sort_by(|a, b| b.0.total_cmp(&a.0));
    let mut cuts: Vec<usize> = gaps.iter().take(k.saturating_sub(1)).filter(|(g, _)| *g > 0.5).map(|(_, i)| *i).collect();
    cuts.sort_unstable();
    let mut out = vec![];
    let mut start = 0;
    for c in cuts.into_iter().chain(std::iter::once(frags.len())) {
        out.push(frags[start..c].to_vec());
        start = c;
    }
    out
}

// ---------- paragraph text ----------

pub(crate) fn shown(r: &Run) -> bool {
    !r.style.vanish
        && !r.tracked_change.as_ref().is_some_and(|c| matches!(c.change_type.as_str(), "delete" | "moveFrom"))
}

pub(crate) fn run_chars(r: &Run) -> impl Iterator<Item = char> + '_ {
    r.text.chars().map(|ch| if matches!(ch, '\x0C' | '\x0B' | '\r') { '\n' } else { ch })
}

/// `orig` of runs `range`, and the code-point offset each run of the
/// paragraph starts at within it (`None` for runs outside or not shown).
pub(crate) fn runs_text(p: &Paragraph, range: std::ops::Range<usize>) -> (String, Vec<Option<usize>>) {
    let mut s = String::new();
    let mut n = 0usize;
    let mut starts = vec![None; p.runs.len()];
    for i in range {
        let r = &p.runs[i];
        if !shown(r) || r.is_math {
            continue;
        }
        starts[i] = Some(n);
        for ch in run_chars(r) {
            s.push(ch);
            n += 1;
        }
    }
    (s, starts)
}

/// Equations of a paragraph, from the document XML when the paragraph could be
/// paired with it, else from Oxi's own expression tree.
type TexMap = HashMap<Vec<usize>, VecDeque<String>>;

fn math_text(r: &Run, tex: Option<&mut VecDeque<String>>) -> String {
    if let Some(t) = tex.and_then(|q| q.pop_front()) {
        return t;
    }
    match &r.style.inline_math {
        Some(mb) => math::block_latex(mb),
        None => r.text.clone(),
    }
}

/// docling's text for a table cell: paragraphs joined with newlines, inline
/// equations as `$…$`, checkbox glyphs dropped.
fn cell_text(blocks: &[Block], tex: &mut TexMap, path: &[usize]) -> String {
    let mut parts = vec![];
    for (k, b) in blocks.iter().enumerate() {
        match b {
            Block::Paragraph(p) => {
                let mut key = path.to_vec();
                key.push(k);
                let mut s = String::new();
                for r in &p.runs {
                    if !shown(r) {
                        continue;
                    }
                    if r.is_math {
                        s.push_str(&format!("${}$", math_text(r, tex.get_mut(&key))));
                    } else {
                        s.extend(run_chars(r).filter(|c| !matches!(c, '☐' | '☒' | '☑')));
                    }
                }
                parts.push(s);
            }
            Block::Math(mb) => {
                let mut key = path.to_vec();
                key.push(k);
                let t = tex.get_mut(&key).and_then(|q| q.pop_front()).unwrap_or_else(|| math::block_latex(mb));
                parts.push(format!("${t}$"));
            }
            Block::Table(t) => {
                for c in t.rows.iter().flat_map(|r| r.cells.iter()) {
                    parts.push(cell_text(&c.blocks, tex, &[]));
                }
            }
            _ => {}
        }
    }
    parts.retain(|s| !s.trim().is_empty());
    parts.join("\n")
}

fn text_provs(frags: &[&Frag], text: &str, len: usize, extra: &[(usize, Rect)]) -> Vec<Value> {
    let mut by_page: BTreeMap<usize, (Rect, Option<(usize, usize)>)> = BTreeMap::new();
    for f in frags {
        let e = by_page.entry(f.page).or_insert((f.rect, None));
        e.0 = e.0.union(f.rect);
    }
    // The characters painted on each page give that page's charspan.
    for (k, loc) in align_locs(text, frags).into_iter().enumerate() {
        if let Some((page, _)) = loc {
            if let Some(e) = by_page.get_mut(&page) {
                e.1 = Some(e.1.map_or((k, k + 1), |(x, y)| (x.min(k), y.max(k + 1))));
            }
        }
    }
    for &(p, r) in extra {
        by_page.entry(p).and_modify(|e| e.0 = e.0.union(r)).or_insert((r, None));
    }
    let many = by_page.len() > 1;
    by_page.into_iter().map(|(page, (rect, span))| {
        let span = match span {
            Some((a, b)) if many => (a.min(len), b.min(len)),
            _ => (0, len),
        };
        prov(page, rect, span)
    }).collect()
}

// ---------- document builder ----------

#[derive(Default)]
struct Builder {
    texts: Vec<Value>,
    tables: Vec<Value>,
    pictures: Vec<Value>,
    groups: Vec<Value>,
    body: Vec<Value>,
    furniture: Vec<Value>,
    list_group: Option<String>,
}

fn rf(s: &str) -> Value {
    json!({"$ref": s})
}

fn obj(v: Value) -> Map<String, Value> {
    match v {
        Value::Object(m) => m,
        _ => unreachable!(),
    }
}

fn formula(tex: &str, provs: Vec<Value>) -> Map<String, Value> {
    obj(json!({"label": "formula", "prov": provs, "orig": tex, "text": tex}))
}

/// Heading level of a paragraph: its own heading/outline level, else its
/// style's (resolved through `basedOn`, as Word and docling do).
pub(crate) fn heading_level(p: &Paragraph, styles: &StyleSheet) -> Option<u8> {
    let st = &p.style;
    let own = st.heading_level.or(st.outline_level.map(|o| o + 1));
    own.or_else(|| {
        let s = styles.styles.get(st.style_id.as_deref()?)?;
        s.paragraph.heading_level.or(s.paragraph.outline_level.map(|o| o + 1))
    }).filter(|l| (1..=9).contains(l))
}

impl Builder {
    fn push(&mut self, kind: &str, parent: &str, mut item: Map<String, Value>) -> String {
        let list = match kind {
            "texts" => &mut self.texts,
            "tables" => &mut self.tables,
            "pictures" => &mut self.pictures,
            _ => &mut self.groups,
        };
        let r = format!("#/{kind}/{}", list.len());
        item.insert("self_ref".into(), json!(r));
        item.insert("parent".into(), rf(parent));
        item.entry("children").or_insert(json!([]));
        item.entry("content_layer").or_insert(json!(if parent == "#/furniture" { "furniture" } else { "body" }));
        list.push(Value::Object(item));
        match parent {
            "#/body" => self.body.push(rf(&r)),
            "#/furniture" => self.furniture.push(rf(&r)),
            g => {
                let i: usize = g.rsplit('/').next().unwrap().parse().unwrap();
                self.groups[i]["children"].as_array_mut().unwrap().push(rf(&r));
            }
        }
        r
    }

    fn group(&mut self, parent: &str, name: &str, label: &str) -> String {
        self.push("groups", parent, obj(json!({"name": name, "label": label})))
    }

    fn picture(&mut self, page: usize, rect: Rect) {
        self.push("pictures", "#/body", obj(json!({"label": "picture", "prov": [prov(page, rect, (0, 0))],
            "captions": [], "references": [], "footnotes": [], "annotations": []})));
    }

    /// A body or text-box paragraph. `frags` are its painted fragments; a text
    /// box paragraph has none and brings its box as `fallback`.
    #[allow(clippy::too_many_arguments)]
    fn paragraph(&mut self, p: &Paragraph, root: &str, frags: &[Frag], markers: &[(usize, Rect)],
                 fallback: Option<(usize, Rect)>, styles: &StyleSheet, tex: Option<&mut VecDeque<String>>) {
        let (mut text, mut starts) = runs_text(p, 0..p.runs.len());
        let has_math = p.runs.iter().any(|r| r.is_math);
        if text.trim().is_empty() && !has_math {
            self.list_group = None;
            return;
        }
        let st = &p.style;
        let style = st.style_id.as_deref().unwrap_or("");
        let heading = heading_level(p, styles);
        let mut marker = st.list_marker.as_deref().filter(|m| !m.trim().is_empty());
        if let (Some(_), Some(m)) = (heading, marker) {
            // docling writes a heading's number into its text, without Word's
            // trailing dot ("1.2 Title").
            let prefix = format!("{} ", m.trim().trim_end_matches('.'));
            let shift = prefix.chars().count();
            text = prefix + &text;
            starts.iter_mut().flatten().for_each(|s| *s += shift);
            marker = None;
        }
        let parent = if marker.is_some() {
            match &self.list_group {
                Some(g) => g.clone(),
                None => {
                    let g = self.group(root, "list", "list");
                    self.list_group = Some(g.clone());
                    g
                }
            }
        } else {
            self.list_group = None;
            root.to_string()
        };

        if has_math {
            return self.math_paragraph(p, &parent, frags, markers, tex);
        }

        let len = text.chars().count();
        let provs = if frags.is_empty() {
            fallback.map(|(pg, r)| vec![prov(pg, r, (0, len))]).unwrap_or_default()
        } else {
            text_provs(&frags.iter().collect::<Vec<_>>(), &text, len, markers)
        };
        let mut item = obj(json!({"prov": provs, "orig": text, "text": text}));
        if let Some(m) = marker {
            item.insert("label".into(), json!("list_item"));
            item.insert("enumerated".into(), json!(m.chars().any(|c| c.is_alphanumeric())));
            item.insert("marker".into(), json!(m));
        } else if style.eq_ignore_ascii_case("title") {
            item.insert("label".into(), json!("title"));
        } else if let Some(l) = heading {
            item.insert("label".into(), json!("section_header"));
            item.insert("level".into(), json!(l));
        } else if style.eq_ignore_ascii_case("caption") {
            item.insert("label".into(), json!("caption"));
        } else {
            item.insert("label".into(), json!("text"));
        }
        self.push("texts", &parent, item);
    }

    /// docling's shape for a paragraph with inline equations: an `inline`
    /// group holding text pieces and `formula` items in reading order.
    fn math_paragraph(&mut self, p: &Paragraph, parent: &str, frags: &[Frag], markers: &[(usize, Rect)],
                      mut tex: Option<&mut VecDeque<String>>) {
        let g = self.group(parent, "group", "inline");
        // Equation glyphs carry no run index; consecutive such fragments form
        // one cluster per inline equation, in reading order.
        let mut clusters: Vec<Vec<(usize, Rect)>> = vec![];
        let mut open = false;
        for f in frags {
            if f.src.is_none() {
                if !open {
                    clusters.push(vec![]);
                    open = true;
                }
                clusters.last_mut().unwrap().push((f.page, f.rect));
            } else {
                open = false;
            }
        }
        // Character position of every run in the paragraph's raw text, which
        // is what a fragment's (run, offset, length) counts in.
        let mut raw_start = vec![0usize];
        for r in &p.runs {
            raw_start.push(raw_start.last().unwrap() + r.text.chars().count());
        }
        let mut cluster_i = 0;
        let mut i = 0;
        let mut first_piece = true;
        while i < p.runs.len() {
            let r = &p.runs[i];
            if r.is_math {
                let t = math_text(r, tex.as_deref_mut());
                let len = t.chars().count();
                // An equation Oxi keeps as plain math text is painted through its
                // run (several adjacent ones may share one fragment: slice it by
                // character share); one it lays out as an expression is painted
                // glyph by glyph without a run index (the next cluster).
                let (a, b) = (raw_start[i], raw_start[i + 1]);
                let own: Vec<(usize, Rect)> = frags.iter().filter_map(|f| {
                    let (run, off, n) = f.src?;
                    let s = raw_start[run] + off;
                    let e = s + n;
                    let (lo, hi) = (s.max(a), e.min(b));
                    (lo < hi && n > 0).then(|| {
                        let w = f.rect.r - f.rect.l;
                        let x = |k: usize| f.rect.l + w * (k - s) as f32 / n as f32;
                        (f.page, Rect { l: x(lo), r: x(hi), ..f.rect })
                    })
                }).collect();
                let provs = if !own.is_empty() {
                    provs_of(own, (0, len))
                } else {
                    let c = clusters.get(cluster_i).map(|c| provs_of(c.iter().copied(), (0, len))).unwrap_or_default();
                    cluster_i += 1;
                    c
                };
                self.push("texts", &g, formula(&t, provs));
                i += 1;
                continue;
            }
            let start = i;
            while i < p.runs.len() && !p.runs[i].is_math {
                i += 1;
            }
            let (text, _) = runs_text(p, start..i);
            if text.trim().is_empty() {
                continue;
            }
            let len = text.chars().count();
            let mine: Vec<&Frag> = frags.iter()
                .filter(|f| f.src.is_some_and(|(run, ..)| (start..i).contains(&run) && !p.runs[run].is_math)).collect();
            let extra = if first_piece { markers } else { &[] };
            first_piece = false;
            let provs = text_provs(&mine, &text, len, extra);
            self.push("texts", &g, obj(json!({"label": "text", "prov": provs, "orig": text, "text": text})));
        }
    }
}

fn table_item(t: &Table, frags: &[Frag], extra: Vec<(usize, Rect)>, tex: &mut TexMap, path: &[usize]) -> Value {
    let mut cells: Vec<Value> = vec![];
    let mut origin_at_col: HashMap<usize, usize> = HashMap::new();
    let mut num_cols = 0usize;
    let mut cell_box: HashMap<(usize, usize), Rect> = HashMap::new();
    let mut first_page: HashMap<(usize, usize), usize> = HashMap::new();
    for f in frags {
        if let Some(k) = f.cell {
            let p = *first_page.entry(k).or_insert(f.page);
            if p == f.page {
                cell_box.entry(k).and_modify(|r| *r = r.union(f.rect)).or_insert(f.rect);
            }
        }
    }
    for (ri, row) in t.rows.iter().enumerate() {
        let mut col = 0usize;
        for (ci, c) in row.cells.iter().enumerate() {
            let span = c.grid_span.max(1) as usize;
            let cont = matches!(c.v_merge.as_deref(), Some(v) if v != "restart");
            if cont {
                if let Some(&i) = origin_at_col.get(&col) {
                    let cell = &mut cells[i];
                    let rs = cell["row_span"].as_u64().unwrap() + 1;
                    cell["row_span"] = json!(rs);
                    cell["end_row_offset_idx"] = json!(ri + 1);
                }
            } else {
                let mut cpath = path.to_vec();
                cpath.extend([ri, ci]);
                let mut cell = json!({
                    "row_span": 1, "col_span": span,
                    "start_row_offset_idx": ri, "end_row_offset_idx": ri + 1,
                    "start_col_offset_idx": col, "end_col_offset_idx": col + span,
                    "text": cell_text(&c.blocks, tex, &cpath),
                    "column_header": ri == 0, "row_header": false, "row_section": false,
                });
                if let Some(r) = cell_box.get(&(ri, ci)) {
                    cell["bbox"] = r.json();
                }
                origin_at_col.insert(col, cells.len());
                cells.push(cell);
            }
            col += span;
        }
        num_cols = num_cols.max(col);
    }
    let provs = provs_of(frags.iter().map(|f| (f.page, f.rect)).chain(extra), (0, 0));
    json!({"label": "table", "prov": provs, "captions": [], "references": [], "footnotes": [],
           "data": {"table_cells": cells, "num_rows": t.rows.len(), "num_cols": num_cols}})
}

/// Where a floating text box sits when the layout painted none of its text
/// (overflow, shapes): its anchor offset from the anchoring paragraph.
pub(crate) fn text_box_rect(doc: &Document, ix: &Index, ci: usize, tb: &TextBox) -> Option<(usize, Rect)> {
    float_rect(doc, ix, ci, tb.anchor_block_index, tb.position.as_ref(), tb.width, tb.height)
}

/// Box of a floating object (text box, shape) from its anchor paragraph's
/// painted position and its offset (`wp:positionH/V`).
pub(crate) fn float_rect(doc: &Document, ix: &Index, ci: usize, anchor_block: usize,
                         position: Option<&oxidocs_core::ir::FloatingPosition>, w: f32, h: f32) -> Option<(usize, Rect)> {
    let sec = doc.pages.get(ci)?;
    let (page, anchor) = ix.anchor_rect((ci, anchor_block), sec.blocks.len())?;
    let (pw, ph) = ix.page_size.get(page - 1).copied()?;
    let (x, y) = match position {
        Some(p) => {
            let x0 = match p.h_relative.as_deref() {
                Some("page") => 0.0,
                Some("character") => anchor.l,
                _ => sec.margin.left,
            };
            let y0 = match p.v_relative.as_deref() {
                Some("page") => 0.0,
                Some("margin") => sec.margin.top,
                _ => anchor.t,
            };
            (x0 + p.x, y0 + p.y)
        }
        None => (anchor.l, anchor.t),
    };
    let x = x.clamp(0.0, pw);
    let y = y.clamp(0.0, ph);
    Some((page, Rect { l: x, t: y, r: (x + w).min(pw), b: (y + h).min(ph) }))
}

fn text_box_group(b: &mut Builder, ix: &mut Index, doc: &Document, ci: usize, tb: &TextBox) {
    let g = b.group("#/body", "textbox", "section");
    let saved = b.list_group.take();
    let fallback = text_box_rect(doc, ix, ci, tb);
    fn walk(b: &mut Builder, ix: &mut Index, doc: &Document, blocks: &[Block], g: &str, fallback: Option<(usize, Rect)>) {
        for blk in blocks {
            match blk {
                Block::Paragraph(p) => {
                    let text = runs_text(p, 0..p.runs.len()).0;
                    let at = ix.take_text(&text).or(fallback);
                    b.paragraph(p, g, &[], &[], at, &doc.styles, None);
                }
                Block::Table(t) => {
                    for c in t.rows.iter().flat_map(|r| r.cells.iter()) {
                        walk(b, ix, doc, &c.blocks, g, fallback);
                    }
                }
                Block::Math(mb) => {
                    let t = math::block_latex(mb);
                    let provs = fallback.map(|(pg, r)| vec![prov(pg, r, (0, t.chars().count()))]).unwrap_or_default();
                    b.push("texts", g, formula(&t, provs));
                }
                _ => {}
            }
        }
    }
    walk(b, ix, doc, &tb.blocks, &g, fallback);
    b.list_group = saved;
}

/// After every item has taken its text: header/footer text that reaches into
/// the body margin (a tall footer) becomes furniture by its position above or
/// below the page's body text, and a list number or bullet whose item started
/// elsewhere (a text box list, a marker on its own line) widens the box of
/// the item on its line.
impl Index {
    /// Header/footer text that reaches into the body margin (a tall footer):
    /// furniture by its position above or below the page's body text. Call
    /// after text boxes have taken their text.
    pub(crate) fn furniture_leftovers(&mut self) {
        let mut body_extent: BTreeMap<usize, (f32, f32)> = BTreeMap::new();
        for f in self.body.values().flatten() {
            let e = body_extent.entry(f.page).or_insert((f.rect.t, f.rect.b));
            e.0 = e.0.min(f.rect.t);
            e.1 = e.1.max(f.rect.b);
        }
        for (&page, evs) in &self.events {
            for ev in evs {
                let Ev::Free(i) = ev else { continue };
                let f = &mut self.free[*i];
                if f.used || !f.has_run {
                    continue;
                }
                if let Some(&(top, bottom)) = body_extent.get(&page) {
                    let header = f.rect.b <= top + 0.5;
                    if header || f.rect.t >= bottom - 0.5 {
                        f.used = true;
                        self.furniture.entry((page, header)).or_default().push((f.text.clone(), f.rect));
                    }
                }
            }
        }
    }
}

fn claim_leftovers(ix: &mut Index, b: &mut Builder) {
    ix.furniture_leftovers();
    let pages: Vec<(usize, Vec<usize>)> = ix.events.iter()
        .map(|(&p, evs)| (p, evs.iter().filter_map(|e| match e { Ev::Free(i) => Some(*i), _ => None }).collect()))
        .collect();
    for (page, frees) in pages {
        for i in frees {
            if ix.free[i].used {
                continue;
            }
            let (rect, has_run) = (ix.free[i].rect, ix.free[i].has_run);
            if has_run {
                continue;
            }
            // A marker: the nearest item to its right on the same line, else
            // the first item starting just below it (the layout can paint a
            // bullet above a first line that grew taller, e.g. for an equation).
            let mid = (rect.t + rect.b) / 2.0;
            let mut same_line: Option<(f32, usize, usize)> = None;
            let mut below: Option<(f32, usize, usize)> = None;
            for (ti, t) in b.texts.iter().enumerate() {
                for (pi, p) in t["prov"].as_array().into_iter().flatten().enumerate() {
                    if p["page_no"].as_u64() != Some(page as u64) {
                        continue;
                    }
                    let bb = &p["bbox"];
                    let (l, tp, bt) = (bb["l"].as_f64().unwrap() as f32, bb["t"].as_f64().unwrap() as f32, bb["b"].as_f64().unwrap() as f32);
                    if l < rect.l - 0.5 {
                        continue;
                    }
                    if tp <= mid && mid <= bt {
                        let d = l - rect.r;
                        if same_line.is_none_or(|(bd, ..)| d < bd) {
                            same_line = Some((d, ti, pi));
                        }
                    } else if tp >= rect.t - 0.5 && tp - rect.b < 2.0 * (rect.b - rect.t) {
                        let d = tp - rect.t;
                        if below.is_none_or(|(bd, ..)| d < bd) {
                            below = Some((d, ti, pi));
                        }
                    }
                }
            }
            if let Some((_, ti, pi)) = same_line.or(below) {
                ix.free[i].used = true;
                let bb = &mut b.texts[ti]["prov"][pi]["bbox"];
                let r2 = |v: f32| ((v as f64) * 100.0).round() / 100.0;
                let l = bb["l"].as_f64().unwrap().min(r2(rect.l));
                let t = bb["t"].as_f64().unwrap().min(r2(rect.t));
                bb["l"] = json!(l);
                bb["t"] = json!(t);
            }
        }
    }
}

pub fn export(doc: &Document, res: &LayoutResult, tex: HashMap<Vec<usize>, Vec<String>>, filename: &str, hash: u64) -> (Value, Stats) {
    let mut ix = index(doc, res);
    let mut tex: TexMap = tex.into_iter().map(|(k, v)| (k, v.into())).collect();
    let mut b = Builder::default();
    let empty: Vec<Frag> = vec![];
    let no_markers: Vec<(usize, Rect)> = vec![];

    for &(page, rect) in ix.pictures.get(&None).into_iter().flatten() {
        b.picture(page, rect);
    }
    // Body blocks are numbered per section in Oxi and once through in the XML.
    let mut global = 0usize;
    // Formula items of a run of consecutive display-equation blocks, laid out
    // up front: no body text separates their glyphs, so they are split by gap.
    let mut pending_math: VecDeque<(String, Vec<Value>)> = VecDeque::new();
    let mut math_shares: VecDeque<usize> = VecDeque::new();
    for (ci, sec) in doc.pages.iter().enumerate() {
        for (bi, block) in sec.blocks.iter().enumerate() {
            let key = (ci, bi);
            let path = vec![global + bi];
            if matches!(block, Block::Math(_)) && pending_math.is_empty() {
                let mut texts: Vec<String> = vec![];
                let mut k = bi;
                while let Some(Block::Math(mb)) = sec.blocks.get(k) {
                    match tex.remove(&vec![global + k]) {
                        Some(q) if !q.is_empty() => texts.extend(q),
                        _ => texts.push(math::block_latex(mb)),
                    }
                    k += 1;
                }
                let groups = clusters(ix.take_between(key), texts.len());
                for (i, t) in texts.into_iter().enumerate() {
                    let provs = groups.get(i).map(|c| provs_of(c.iter().copied(), (0, t.chars().count()))).unwrap_or_default();
                    pending_math.push_back((t, provs));
                }
                // Each block of the run emits its own share below, in order.
                let mut per_block: Vec<usize> = vec![];
                let mut k = bi;
                while let Some(Block::Math(_)) = sec.blocks.get(k) {
                    per_block.push(1);
                    k += 1;
                }
                let extra = pending_math.len().saturating_sub(per_block.len());
                if let Some(first) = per_block.first_mut() {
                    *first += extra; // several oMath in one paragraph: all with the first block
                }
                math_shares = per_block.into();
            }
            match block {
                Block::Paragraph(p) => {
                    let frags = ix.body.get(&key).unwrap_or(&empty);
                    let markers = ix.markers.get(&key).unwrap_or(&no_markers).clone();
                    b.paragraph(p, "#/body", frags, &markers, None, &doc.styles, tex.get_mut(&path));
                }
                Block::Table(t) => {
                    b.list_group = None;
                    let extra = ix.take_between(key);
                    let frags = ix.body.get(&key).unwrap_or(&empty);
                    b.push("tables", "#/body", obj(table_item(t, frags, extra, &mut tex, &path)));
                }
                Block::Math(_) => {
                    b.list_group = None;
                    let n = math_shares.pop_front().unwrap_or(1);
                    for _ in 0..n {
                        if let Some((t, provs)) = pending_math.pop_front() {
                            b.push("texts", "#/body", formula(&t, provs));
                        }
                    }
                }
                _ => {}
            }
            for &(page, rect) in ix.pictures.get(&Some(key)).into_iter().flatten() {
                b.picture(page, rect);
            }
            for tb in sec.text_boxes.iter().filter(|t| t.anchor_block_index == bi) {
                text_box_group(&mut b, &mut ix, doc, ci, tb);
            }
        }
        global += sec.blocks.len();
    }

    claim_leftovers(&mut ix, &mut b);
    for (&(page, header), frs) in &ix.furniture {
        let text = frs.iter().map(|(t, _)| t.as_str()).collect::<String>().trim().to_string();
        if text.is_empty() {
            continue;
        }
        let rect = frs.iter().map(|(_, r)| *r).reduce(Rect::union).unwrap();
        let len = text.chars().count();
        let item = obj(json!({"label": if header { "page_header" } else { "page_footer" },
            "prov": [prov(page, rect, (0, len))], "orig": text, "text": text}));
        b.push("texts", "#/furniture", item);
    }

    let mut pages = Map::new();
    for (i, pg) in res.pages.iter().enumerate() {
        pages.insert((i + 1).to_string(), json!({"size": {"width": pg.width, "height": pg.height}, "page_no": i + 1}));
    }
    if std::env::var_os("OXI_DOCLING_DEBUG").is_some() {
        for (page, evs) in &ix.events {
            for ev in evs {
                if let Ev::Free(i) = ev {
                    let f = &ix.free[*i];
                    if !f.used {
                        eprintln!("[unplaced] p{page} run={} xy {:.0},{:.0} {:?}", f.has_run, f.rect.l, f.rect.t, f.text);
                    }
                }
            }
        }
    }
    let all: Vec<&Value> = b.texts.iter().chain(&b.tables).chain(&b.pictures).collect();
    let stats = Stats {
        pages: res.pages.len(),
        items: all.len(),
        items_with_prov: all.iter().filter(|v| v["prov"].as_array().is_some_and(|p| !p.is_empty())).count(),
        unplaced_fragments: ix.free.iter().filter(|f| !f.used).count(),
    };
    let stem = filename.rsplit(['/', '\\']).next().unwrap_or(filename);
    let name = stem.rsplit_once('.').map_or(stem, |(s, _)| s);
    let out = json!({
        "schema_name": "DoclingDocument",
        "version": SCHEMA_VERSION,
        "name": name,
        "origin": {"mimetype": DOCX_MIME, "binary_hash": hash, "filename": stem},
        "furniture": {"self_ref": "#/furniture", "children": b.furniture, "content_layer": "furniture", "name": "_root_", "label": "unspecified"},
        "body": {"self_ref": "#/body", "children": b.body, "content_layer": "body", "name": "_root_", "label": "unspecified"},
        "groups": b.groups, "texts": b.texts, "pictures": b.pictures, "tables": b.tables,
        "key_value_items": [], "form_items": [], "pages": pages,
    });
    (out, stats)
}

pub(crate) fn fnv1a(data: &[u8]) -> u64 {
    let mut h: u64 = 0xcbf29ce484222325;
    for &x in data {
        h ^= x as u64;
        h = h.wrapping_mul(0x100000001b3);
    }
    h
}
