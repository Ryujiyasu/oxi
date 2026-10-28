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
//! - A text box is the union of the item's line boxes on that page.
//! - `charspan` indexes the item's `orig` in Unicode code points (what Python
//!   `str` slicing uses). `orig` is the concatenation of the paragraph's runs as
//!   Word shows them in its final view: deleted revisions and hidden text are
//!   left out. `text` equals `orig`.
//! - A table's and a picture's `charspan` is `[0, 0]`.
//!
//! Structure follows docling's own docx output where it matters for consumers:
//! headers/footers are `page_header` / `page_footer` furniture (one item per
//! page), a paragraph with inline equations becomes an `inline` group of text
//! and `formula` items, a display equation a `formula` item, a text box a
//! `textbox` group placed after the paragraph it anchors to. Footnotes and
//! text inside shapes are not exported yet; [`Stats::unplaced_fragments`]
//! counts the text the layout painted that found no item.
//!
//! [DoclingDocument]: https://github.com/docling-project/docling-core

mod math;

use std::collections::{BTreeMap, HashMap};

use oxidocs_core::ir::{Block, Document, Paragraph, Run, ShowRevisions, Table, TextBox};
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
    Ok(export(&doc, &res, filename, fnv1a(data)))
}

// ---------- geometry ----------

#[derive(Clone, Copy)]
struct Rect {
    l: f32,
    t: f32,
    r: f32,
    b: f32,
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

fn prov(page: usize, rect: Rect, span: (usize, usize)) -> Value {
    json!({"page_no": page, "bbox": rect.json(), "charspan": [span.0, span.1]})
}

/// Union of (page, rect) pairs per page, in page order.
fn union_by_page(it: impl IntoIterator<Item = (usize, Rect)>) -> BTreeMap<usize, Rect> {
    let mut m: BTreeMap<usize, Rect> = BTreeMap::new();
    for (p, r) in it {
        m.entry(p).and_modify(|u| *u = u.union(r)).or_insert(r);
    }
    m
}

// ---------- the layout, indexed ----------

type Key = (usize, usize); // (section, body block)

/// One painted text fragment of a body block.
struct Frag {
    page: usize,
    rect: Rect,
    /// (run index, char offset in run, chars) when the fragment maps to source
    /// text; `None` for inline equations, which the layout paints per glyph.
    src: Option<(usize, usize, usize)>,
    /// Top-level table cell (row, cell index in row) for table fragments.
    cell: Option<(usize, usize)>,
}

/// A painted text fragment with no body block: display equations, text boxes,
/// list bullets, footnotes.
struct Free {
    text: String,
    rect: Rect,
    /// Text boxes keep their run index; equation glyphs and bullets have none.
    has_run: bool,
    used: bool,
}

enum Ev {
    Body(Key),
    Free(usize), // index into Index::free
}

#[derive(Default)]
struct Index {
    body: HashMap<Key, Vec<Frag>>,
    /// Pictures, keyed by the body block painted just before them.
    pictures: HashMap<Option<Key>, Vec<(usize, Rect)>>,
    /// (page, is_header) -> (text, rect) in paint order.
    furniture: BTreeMap<(usize, bool), Vec<(String, Rect)>>,
    free: Vec<Free>,
    /// Paint order per page (1-based page -> events).
    events: BTreeMap<usize, Vec<Ev>>,
}

fn index(doc: &Document, res: &LayoutResult) -> Index {
    let mut ix = Index::default();
    let mut last_body: Option<Key> = None;
    for (pi, pg) in res.pages.iter().enumerate() {
        let page = pi + 1;
        let sec = pg.elements.iter().find_map(|e| e.source_container_index.filter(|_| e.paragraph_index.is_some()))
            .or(last_body.map(|k| k.0)).unwrap_or(0);
        let margin = doc.pages.get(sec).map(|s| (s.margin.top, s.margin.bottom)).unwrap_or((72.0, 72.0));
        let in_header = |r: &Rect| r.b <= margin.0 + 1.0;
        let in_footer = |r: &Rect| r.t >= pg.height - margin.1 - 1.0;
        let evs = ix.events.entry(page).or_default();
        for e in &pg.elements {
            match &e.content {
                LayoutContent::Text { text, .. } => {
                    let top = e.y - e.text_y_off;
                    let rect = Rect { l: e.x, t: top, r: e.x + e.width, b: top + e.height };
                    match (e.source_container_index, e.paragraph_index) {
                        (Some(c), Some(p)) => {
                            let src = match (e.run_index, e.char_offset) {
                                (Some(r), Some(o)) => Some((r, o, e.source_char_len.unwrap_or_else(|| text.chars().count()))),
                                _ => None,
                            };
                            let cell = match e.cell_ancestor_path.first() {
                                Some(&(row, cell, _)) => Some((row, cell)),
                                None => e.cell_row_index.zip(e.cell_col_index),
                            };
                            ix.body.entry((c, p)).or_default().push(Frag { page, rect, src, cell });
                            evs.push(Ev::Body((c, p)));
                            last_body = Some((c, p));
                        }
                        _ if text.trim().is_empty() => {}
                        _ if in_header(&rect) => ix.furniture.entry((page, true)).or_default().push((text.clone(), rect)),
                        _ if in_footer(&rect) => ix.furniture.entry((page, false)).or_default().push((text.clone(), rect)),
                        _ => {
                            evs.push(Ev::Free(ix.free.len()));
                            ix.free.push(Free { text: text.clone(), rect, has_run: e.run_index.is_some(), used: false });
                        }
                    }
                }
                LayoutContent::Image { .. } => {
                    let rect = Rect { l: e.x, t: e.y, r: e.x + e.width, b: e.y + e.height };
                    if !in_header(&rect) && !in_footer(&rect) {
                        ix.pictures.entry(last_body).or_default().push((page, rect));
                    }
                }
                _ => {}
            }
        }
    }
    ix
}

fn norm(s: &str) -> String {
    s.chars().filter(|c| !c.is_whitespace()).collect()
}

fn is_bullet(s: &str) -> bool {
    let mut it = s.trim().chars();
    match (it.next(), it.next()) {
        (Some(c), None) => matches!(c, '\u{E000}'..='\u{F8FF}' | '•' | '◦' | '▪' | '■' | '●' | '○' | '–' | '-' | '*'),
        _ => false,
    }
}

impl Index {
    /// Free fragments painted between this block's neighbours, on every page:
    /// where a display equation's glyphs land (they carry no block index).
    fn take_between(&mut self, key: Key) -> Vec<(usize, Rect)> {
        let mut out = vec![];
        for (&page, evs) in &self.events {
            let before = evs.iter().rposition(|e| matches!(e, Ev::Body(k) if *k < key));
            let after = evs.iter().position(|e| matches!(e, Ev::Body(k) if *k > key));
            let lo = before.map_or(0, |i| i + 1);
            let hi = after.unwrap_or(evs.len());
            if before.is_none() && after.is_none() && !evs.is_empty() {
                continue; // a page with no body text at all: not this block's
            }
            for ev in &evs[lo..hi] {
                if let Ev::Free(i) = ev {
                    let f = &mut self.free[*i];
                    if !f.used && !f.has_run && !is_bullet(&f.text) {
                        f.used = true;
                        out.push((page, f.rect));
                    }
                }
            }
        }
        out
    }

    /// Consecutive free fragments (with a run index, in paint order on one
    /// page) whose text concatenates to `target`: a text box paragraph.
    fn take_text(&mut self, target: &str) -> Option<(usize, Rect)> {
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
                for &i in &idx[start..] {
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
                        for &j in &idx[start..=idx.iter().position(|&x| x == i).unwrap()] {
                            self.free[j].used = true;
                        }
                        return Some((page, rect.unwrap()));
                    }
                }
            }
        }
        None
    }
}

// ---------- paragraph text ----------

fn shown(r: &Run) -> bool {
    !r.style.vanish
        && !r.tracked_change.as_ref().is_some_and(|c| matches!(c.change_type.as_str(), "delete" | "moveFrom"))
}

fn run_chars(r: &Run) -> impl Iterator<Item = char> + '_ {
    r.text.chars().map(|ch| if matches!(ch, '\x0C' | '\x0B' | '\r') { '\n' } else { ch })
}

/// `orig` of runs `range`, and the code-point offset each run of the
/// paragraph starts at within it (`None` for runs outside or not shown).
fn runs_text(p: &Paragraph, range: std::ops::Range<usize>) -> (String, Vec<Option<usize>>) {
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

fn blocks_text(blocks: &[Block]) -> String {
    let mut parts = vec![];
    for b in blocks {
        match b {
            Block::Paragraph(p) => parts.push(runs_text(p, 0..p.runs.len()).0),
            Block::Table(t) => parts.extend(t.rows.iter().flat_map(|r| r.cells.iter()).map(|c| blocks_text(&c.blocks))),
            _ => {}
        }
    }
    parts.retain(|s| !s.trim().is_empty());
    parts.join(" ")
}

/// Provenance of a text item from the fragments of its runs: one entry per
/// page, with the char range painted there.
fn text_provs(frags: &[&Frag], starts: &[Option<usize>], len: usize) -> Vec<Value> {
    let mut by_page: BTreeMap<usize, (Rect, Option<(usize, usize)>)> = BTreeMap::new();
    for f in frags {
        let span = f.src.and_then(|(run, off, n)| starts.get(run).copied().flatten().map(|s| (s + off, s + off + n)));
        let e = by_page.entry(f.page).or_insert((f.rect, span));
        e.0 = e.0.union(f.rect);
        if let Some((a, b)) = span {
            e.1 = Some(e.1.map_or((a, b), |(x, y)| (x.min(a), y.max(b))));
        }
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

    /// A body or text-box paragraph. `frags` are its painted fragments (empty
    /// when the caller resolves provenance itself through `fallback`).
    fn paragraph(&mut self, p: &Paragraph, root: &str, frags: &[Frag], fallback: Option<(usize, Rect)>) {
        let (mut text, mut starts) = runs_text(p, 0..p.runs.len());
        let has_math = p.runs.iter().any(|r| r.is_math && r.style.inline_math.is_some());
        if text.trim().is_empty() && !has_math {
            self.list_group = None;
            return;
        }
        let st = &p.style;
        let style = st.style_id.as_deref().unwrap_or("");
        let heading = st.heading_level.or(st.outline_level.map(|o| o + 1)).filter(|l| *l <= 9);
        let mut marker = st.list_marker.as_deref().filter(|m| !m.trim().is_empty());
        if let (Some(_), Some(m)) = (heading, marker) {
            // A numbered heading is a heading. docling writes the number into
            // the text without Word's trailing dot ("1.2 Title"); do the same
            // and shift the run offsets.
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
            return self.math_paragraph(p, &parent, frags);
        }

        let len = text.chars().count();
        let provs = if frags.is_empty() {
            fallback.map(|(pg, r)| vec![prov(pg, r, (0, len))]).unwrap_or_default()
        } else {
            text_provs(&frags.iter().collect::<Vec<_>>(), &starts, len)
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
            item.insert("level".into(), json!(l.max(1)));
        } else if style.eq_ignore_ascii_case("caption") {
            item.insert("label".into(), json!("caption"));
        } else {
            item.insert("label".into(), json!("text"));
        }
        self.push("texts", &parent, item);
    }

    /// docling's shape for a paragraph with inline equations: an `inline`
    /// group holding text pieces and `formula` items in reading order.
    fn math_paragraph(&mut self, p: &Paragraph, parent: &str, frags: &[Frag]) {
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
        let mut cluster_i = 0;
        let mut i = 0;
        while i < p.runs.len() {
            let r = &p.runs[i];
            if r.is_math {
                if let Some(mb) = &r.style.inline_math {
                    let tex = math::block_latex(mb);
                    let provs: Vec<Value> = clusters.get(cluster_i).map(|c| {
                        union_by_page(c.iter().copied()).into_iter()
                            .map(|(pg, rc)| prov(pg, rc, (0, tex.chars().count()))).collect()
                    }).unwrap_or_default();
                    cluster_i += 1;
                    self.push("texts", &g, obj(json!({"label": "formula", "prov": provs, "orig": tex, "text": tex})));
                }
                i += 1;
                continue;
            }
            let start = i;
            while i < p.runs.len() && !p.runs[i].is_math {
                i += 1;
            }
            let (text, starts) = runs_text(p, start..i);
            if text.trim().is_empty() {
                continue;
            }
            let len = text.chars().count();
            let mine: Vec<&Frag> = frags.iter().filter(|f| f.src.is_some_and(|(run, ..)| (start..i).contains(&run))).collect();
            let provs = text_provs(&mine, &starts, len);
            self.push("texts", &g, obj(json!({"label": "text", "prov": provs, "orig": text, "text": text})));
        }
    }
}

fn table_item(t: &Table, frags: &[Frag]) -> Value {
    let mut cells: Vec<Value> = vec![];
    let mut origin_at_col: HashMap<usize, usize> = HashMap::new();
    let mut num_cols = 0usize;
    // A cell's box: its fragments on the first page it appears on.
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
                let mut cell = json!({
                    "row_span": 1, "col_span": span,
                    "start_row_offset_idx": ri, "end_row_offset_idx": ri + 1,
                    "start_col_offset_idx": col, "end_col_offset_idx": col + span,
                    "text": blocks_text(&c.blocks),
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
    let provs: Vec<Value> = union_by_page(frags.iter().map(|f| (f.page, f.rect))).into_iter()
        .map(|(pg, r)| prov(pg, r, (0, 0))).collect();
    json!({"label": "table", "prov": provs, "captions": [], "references": [], "footnotes": [],
           "data": {"table_cells": cells, "num_rows": t.rows.len(), "num_cols": num_cols}})
}

fn text_box_group(b: &mut Builder, ix: &mut Index, tb: &TextBox) {
    let g = b.group("#/body", "textbox", "section");
    let saved = b.list_group.take();
    fn walk(b: &mut Builder, ix: &mut Index, blocks: &[Block], g: &str) {
        for blk in blocks {
            match blk {
                Block::Paragraph(p) => {
                    let text = runs_text(p, 0..p.runs.len()).0;
                    let at = ix.take_text(&text);
                    b.paragraph(p, g, &[], at);
                }
                Block::Table(t) => {
                    for c in t.rows.iter().flat_map(|r| r.cells.iter()) {
                        walk(b, ix, &c.blocks, g);
                    }
                }
                _ => {}
            }
        }
    }
    walk(b, ix, &tb.blocks, &g);
    b.list_group = saved;
}

pub fn export(doc: &Document, res: &LayoutResult, filename: &str, hash: u64) -> (Value, Stats) {
    let mut ix = index(doc, res);
    let mut b = Builder::default();
    let empty: Vec<Frag> = vec![];

    for &(page, rect) in ix.pictures.get(&None).into_iter().flatten() {
        b.picture(page, rect);
    }
    for (ci, sec) in doc.pages.iter().enumerate() {
        for (bi, block) in sec.blocks.iter().enumerate() {
            let key = (ci, bi);
            match block {
                Block::Paragraph(p) => {
                    let frags = ix.body.get(&key).unwrap_or(&empty);
                    b.paragraph(p, "#/body", frags, None);
                }
                Block::Table(t) => {
                    b.list_group = None;
                    let frags = ix.body.get(&key).unwrap_or(&empty);
                    b.push("tables", "#/body", obj(table_item(t, frags)));
                }
                Block::Math(mb) => {
                    b.list_group = None;
                    let tex = math::block_latex(mb);
                    let len = tex.chars().count();
                    let provs: Vec<Value> = union_by_page(ix.take_between(key)).into_iter()
                        .map(|(pg, r)| prov(pg, r, (0, len))).collect();
                    b.push("texts", "#/body", obj(json!({"label": "formula", "prov": provs, "orig": tex, "text": tex})));
                }
                _ => {}
            }
            for &(page, rect) in ix.pictures.get(&Some(key)).into_iter().flatten() {
                b.picture(page, rect);
            }
            for tb in sec.text_boxes.iter().filter(|t| t.anchor_block_index == bi) {
                text_box_group(&mut b, &mut ix, tb);
            }
        }
    }

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
    let all: Vec<&Value> = b.texts.iter().chain(&b.tables).chain(&b.pictures).collect();
    let stats = Stats {
        pages: res.pages.len(),
        items: all.len(),
        items_with_prov: all.iter().filter(|v| v["prov"].as_array().is_some_and(|p| !p.is_empty())).count(),
        unplaced_fragments: ix.free.iter().filter(|f| !f.used && !is_bullet(&f.text)).count(),
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

fn fnv1a(data: &[u8]) -> u64 {
    let mut h: u64 = 0xcbf29ce484222325;
    for &x in data {
        h ^= x as u64;
        h = h.wrapping_mul(0x100000001b3);
    }
    h
}
