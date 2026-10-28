// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! Annotate mode: take the DoclingDocument that docling itself produced for a
//! .docx and fill every item's `prov` (and the document's `pages`) from Oxi's
//! layout of the same file. docling's structure, labels and text are kept
//! exactly; only `prov` and `pages` are written, plus a table cell's `bbox`.
//!
//! Oxi's painted text is turned into one stream of located characters in
//! reading order (body paragraphs, table cells row by row, text boxes after
//! their anchor). Each docling text item is found in that stream, searching
//! forward from where the previous item ended, with whitespace ignored. The
//! characters it covers give one box per page; `charspan` is the range of the
//! item's `orig` painted on that page. Formulas and pictures are matched in
//! order, tables through their cells' text, headers/footers per page, and an
//! item with no text of its own (a list item holding an inline group) takes
//! the union of its children.

use std::collections::{BTreeMap, HashMap};

use oxidocs_core::ir::{Block, Document, Paragraph, ShowRevisions};
use oxidocs_core::layout::{LayoutEngine, LayoutResult};
use serde_json::{json, Value};

use crate::{clusters, heading_level, index, prov, runs_text, union_by_page, Frag, Index, Rect};

type Loc = (usize, Rect);

#[derive(Debug, Default, Clone)]
pub struct AnnotateStats {
    pub pages: usize,
    pub items: usize,
    pub items_with_prov: usize,
    /// Items that should carry provenance but could not be placed: (ref, text).
    pub unmatched: Vec<(String, String)>,
}

fn ignorable(c: char) -> bool {
    c.is_whitespace() || matches!(c, '☐' | '☒' | '☑')
}

/// The located-character stream plus the ordered formula and picture places.
#[derive(Default)]
struct Painted {
    chars: Vec<char>,
    locs: Vec<Option<Loc>>,
    /// List number/bullet boxes of a paragraph, keyed by the stream index of
    /// its first character: added to an item that starts there.
    markers: HashMap<usize, Vec<Loc>>,
    formulas: Vec<Vec<Loc>>,
    /// Picture places: (stream position, box).
    pictures: Vec<(usize, Loc)>,
    /// Empty paragraphs: (stream position, box Oxi gave the empty line).
    empties: Vec<(usize, Option<Loc>)>,
    /// Every paragraph's stream start and first located character.
    para_starts: Vec<(usize, Loc)>,
}

impl Painted {
    /// A paragraph with no visible text: Oxi still lays out its empty line.
    fn push_empty_if_blank(&mut self, text: &str, p: &Paragraph, frags: &[&Frag]) -> bool {
        if text.chars().any(|c| !ignorable(c)) || p.runs.iter().any(|r| r.is_math) {
            return false;
        }
        let at = union_by_page(frags.iter().map(|f| (f.page, f.rect))).into_iter().next();
        self.empties.push((self.chars.len(), at));
        true
    }

    fn push_text(&mut self, text: &str, locs: &[Option<Loc>]) -> usize {
        let start = self.chars.len();
        if let Some(first) = locs.iter().flatten().next() {
            self.para_starts.push((start, *first));
        }
        for (i, c) in text.chars().enumerate() {
            if !ignorable(c) {
                self.chars.push(c);
                self.locs.push(locs.get(i).copied().flatten());
            }
        }
        start
    }
}

/// Per-character locations of a paragraph's `orig`, from its fragments.
fn char_locs(p: &Paragraph, frags: &[&Frag]) -> (String, Vec<Option<Loc>>) {
    let (text, _) = runs_text(p, 0..p.runs.len());
    let locs = crate::align_locs(&text, frags);
    (text, locs)
}

/// Formula places of a paragraph's math runs, in order (see `math_paragraph`).
fn para_formulas(p: &Paragraph, frags: &[&Frag]) -> Vec<Vec<Loc>> {
    let mut raw_start = vec![0usize];
    for r in &p.runs {
        raw_start.push(raw_start.last().unwrap() + r.text.chars().count());
    }
    let mut glyph_clusters: Vec<Vec<Loc>> = vec![];
    let mut open = false;
    for f in frags {
        if f.src.is_none() {
            if !open {
                glyph_clusters.push(vec![]);
                open = true;
            }
            glyph_clusters.last_mut().unwrap().push((f.page, f.rect));
        } else {
            open = false;
        }
    }
    let mut out = vec![];
    let mut ci = 0;
    for (i, r) in p.runs.iter().enumerate() {
        if !r.is_math {
            continue;
        }
        let (a, b) = (raw_start[i], raw_start[i + 1]);
        let own: Vec<Loc> = frags.iter().filter_map(|f| {
            let (run, off, n) = f.src?;
            let s = raw_start[run] + off;
            let (lo, hi) = (s.max(a), (s + n).min(b));
            (lo < hi && n > 0).then(|| {
                let w = f.rect.r - f.rect.l;
                let x = |k: usize| f.rect.l + w * (k - s) as f32 / n as f32;
                (f.page, Rect { l: x(lo), r: x(hi), ..f.rect })
            })
        }).collect();
        if !own.is_empty() {
            out.push(own);
        } else {
            out.push(glyph_clusters.get(ci).cloned().unwrap_or_default());
            ci += 1;
        }
    }
    out
}

fn flat_paragraphs<'a>(blocks: &'a [Block], out: &mut Vec<&'a Paragraph>) {
    for b in blocks {
        match b {
            Block::Paragraph(p) => out.push(p),
            Block::Table(t) => t.rows.iter().flat_map(|r| r.cells.iter()).for_each(|c| flat_paragraphs(&c.blocks, out)),
            _ => {}
        }
    }
}

/// After a body block: its pictures at the current reading position (each
/// painted bitmap, plus one box for the blip-less shapes and text box frames
/// anchored to it, which docling turns into one picture per paragraph), then
/// the text of the text boxes anchored to it.
fn after_block(pt: &mut Painted, ix: &mut Index, doc: &Document, ci: usize, bi: usize) {
    let sec = &doc.pages[ci];
    let pos = pt.chars.len();
    for &l in ix.pictures.get(&Some((ci, bi))).into_iter().flatten() {
        pt.pictures.push((pos, l));
    }
    let shapes = match sec.blocks.get(bi) {
        Some(Block::Paragraph(p)) => p.shapes.iter()
            .filter_map(|s| crate::float_rect(doc, ix, ci, bi, s.position.as_ref(), s.width, s.height)).collect(),
        _ => vec![],
    };
    let frames: Vec<Loc> = sec.text_boxes.iter().filter(|t| t.anchor_block_index == bi)
        .filter_map(|t| crate::text_box_rect(doc, ix, ci, t)).collect();
    if let Some((page, r)) = union_by_page(shapes.into_iter().chain(frames)).into_iter().next() {
        pt.pictures.push((pos, (page, r)));
    }
    for tb in sec.text_boxes.iter().filter(|t| t.anchor_block_index == bi) {
        text_box(pt, ix, doc, ci, tb);
    }
}

fn build(doc: &Document, ix: &mut Index) -> Painted {
    let mut pt = Painted::default();
    let empty: Vec<Frag> = vec![];
    for &l in ix.pictures.get(&None).into_iter().flatten() {
        pt.pictures.push((0, l));
    }
    for (ci, sec) in doc.pages.iter().enumerate() {
        let mut bi = 0;
        while bi < sec.blocks.len() {
            let key = (ci, bi);
            let frags: Vec<&Frag> = ix.body.get(&key).unwrap_or(&empty).iter().collect();
            match &sec.blocks[bi] {
                Block::Paragraph(p) => {
                    let (text, locs) = char_locs(p, &frags);
                    if !pt.push_empty_if_blank(&text, p, &frags) {
                        let marker = p.style.list_marker.as_deref().map(str::trim).filter(|m| !m.is_empty());
                        let markers = ix.markers.get(&key).cloned().unwrap_or_default();
                        let start = pt.chars.len();
                        if let (Some(_), Some(m)) = (heading_level(p, &doc.styles), marker) {
                            // docling writes a heading's number into its text.
                            let num = m.trim_end_matches('.');
                            let at = markers.first().copied().or_else(|| locs.iter().flatten().next().copied());
                            pt.push_text(num, &vec![at; num.chars().count()]);
                        }
                        pt.push_text(&text, &locs);
                        if !markers.is_empty() {
                            pt.markers.insert(start, markers);
                        }
                        pt.formulas.extend(para_formulas(p, &frags));
                    }
                    after_block(&mut pt, ix, doc, ci, bi);
                    bi += 1;
                }
                Block::Table(t) => {
                    for (ri, row) in t.rows.iter().enumerate() {
                        for (cj, cell) in row.cells.iter().enumerate() {
                            let mut paras = vec![];
                            flat_paragraphs(&cell.blocks, &mut paras);
                            for (k, p) in paras.iter().enumerate() {
                                let mine: Vec<&Frag> = frags.iter().copied()
                                    .filter(|f| f.cell == Some((ri, cj)) && f.cpi.is_none_or(|c| c == k)).collect();
                                let (text, mut locs) = char_locs(p, &mine);
                                if pt.push_empty_if_blank(&text, p, &mine) {
                                    continue;
                                }
                                // A nested paragraph whose fragments could not be told
                                // apart takes its cell's box.
                                if locs.iter().all(Option::is_none) {
                                    let cell_box = union_by_page(mine.iter().map(|f| (f.page, f.rect))).into_iter().next();
                                    locs = vec![cell_box; locs.len()];
                                }
                                pt.push_text(&text, &locs);
                                pt.formulas.extend(para_formulas(p, &mine));
                            }
                        }
                    }
                    after_block(&mut pt, ix, doc, ci, bi);
                    bi += 1;
                }
                Block::Math(_) => {
                    // A run of display equations: their glyphs, split by gap.
                    let mut n = 0;
                    while matches!(sec.blocks.get(bi + n), Some(Block::Math(_))) {
                        n += 1;
                    }
                    let groups = clusters(ix.take_between(key), n);
                    for g in 0..n {
                        pt.formulas.push(groups.get(g).cloned().unwrap_or_default());
                    }
                    for k in bi..bi + n {
                        after_block(&mut pt, ix, doc, ci, k);
                    }
                    bi += n;
                }
                _ => {
                    after_block(&mut pt, ix, doc, ci, bi);
                    bi += 1;
                }
            }
        }
    }
    pt
}

fn text_box(pt: &mut Painted, ix: &mut Index, doc: &Document, ci: usize, tb: &oxidocs_core::ir::TextBox) {
    let fallback = crate::text_box_rect(doc, ix, ci, tb);
    let mut paras = vec![];
    flat_paragraphs(&tb.blocks, &mut paras);
    for p in paras {
        let (text, _) = runs_text(p, 0..p.runs.len());
        if !text.chars().any(|c| !ignorable(c)) {
            pt.empties.push((pt.chars.len(), fallback));
            continue;
        }
        let at = ix.take_text(&text).or(fallback);
        pt.push_text(&text, &vec![at; text.chars().count()]);
    }
}

// ---------- matching ----------

struct Matcher<'a> {
    pt: &'a Painted,
    hay: String,
    /// byte offset in `hay` -> char index, for every char start.
    char_at: HashMap<usize, usize>,
    byte_of: Vec<usize>,
    cursor: usize,
    formula_i: usize,
    empty_used: Vec<bool>,
    picture_used: Vec<bool>,
}

impl<'a> Matcher<'a> {
    fn new(pt: &'a Painted) -> Self {
        let hay: String = pt.chars.iter().collect();
        let mut char_at = HashMap::new();
        let mut byte_of = vec![];
        for (ci, (b, _)) in hay.char_indices().enumerate() {
            char_at.insert(b, ci);
            byte_of.push(b);
        }
        byte_of.push(hay.len());
        let empty_used = vec![false; pt.empties.len()];
        let picture_used = vec![false; pt.pictures.len()];
        Matcher { pt, hay, char_at, byte_of, cursor: 0, formula_i: 0, empty_used, picture_used }
    }

    /// The first unused picture place at or after the reading cursor, else
    /// any unused one, else the nearest place (two docling pictures can come
    /// from one paragraph's drawings).
    fn picture(&mut self) -> Option<Loc> {
        let pics = &self.pt.pictures;
        let i = (0..pics.len()).find(|&i| !self.picture_used[i] && pics[i].0 >= self.cursor)
            .or_else(|| (0..pics.len()).find(|&i| !self.picture_used[i]))
            .or_else(|| (0..pics.len()).min_by_key(|&i| pics[i].0.abs_diff(self.cursor)))?;
        self.picture_used[i] = true;
        Some(pics[i].1)
    }

    /// The next empty paragraph at or after the cursor: the empty line Oxi
    /// laid out for it. One Oxi paints nothing for (a section-break mark) or
    /// cannot pair takes a zero-height box at the end of the text before it.
    fn empty(&mut self) -> Option<Vec<Value>> {
        let found = (0..self.pt.empties.len()).find(|&i| !self.empty_used[i] && self.pt.empties[i].0 >= self.cursor);
        let (pos, at) = match found {
            Some(i) => {
                self.empty_used[i] = true;
                self.pt.empties[i]
            }
            None => (self.cursor, None),
        };
        self.cursor = self.cursor.max(pos);
        let at = at.or_else(|| {
            (0..pos.min(self.pt.locs.len())).rev().find_map(|k| self.pt.locs[k])
                .map(|(page, r)| (page, Rect { t: r.b, b: r.b, ..r }))
        });
        at.map(|(page, r)| vec![prov(page, r, (0, 0))])
    }

    /// Stream range of `needle`: forward from the cursor, else anywhere.
    fn find(&self, needle: &str, from: usize) -> Option<(usize, usize)> {
        if needle.is_empty() {
            return None;
        }
        let n = needle.chars().count();
        let hit = |start_byte: usize| self.hay[start_byte..].find(needle).map(|b| self.char_at[&(start_byte + b)]);
        let s = hit(self.byte_of[from.min(self.byte_of.len() - 1)]).or_else(|| hit(0))?;
        Some((s, s + n))
    }

    /// Locate a text item; returns its prov entries.
    fn text(&mut self, orig: &str) -> Option<Vec<Value>> {
        // Positions in `orig` of the characters the stream keeps.
        let kept: Vec<(usize, char)> = orig.chars().enumerate().filter(|(_, c)| !ignorable(*c)).collect();
        let full: String = kept.iter().map(|(_, c)| c).collect();
        let mut tries: Vec<(String, usize)> = vec![(full.clone(), 0)];
        // A heading number docling counted differently from Word: match the title alone.
        let num_len = full.chars().take_while(|c| c.is_ascii_digit() || *c == '.').count();
        if num_len > 0 && num_len < full.chars().count() {
            tries.push((full.chars().skip(num_len).collect(), num_len));
        }
        // Inline `$…$` equations in a cell's text are not painted as text.
        if full.contains('$') {
            let stripped: String = full.split('$').step_by(2).collect();
            tries.push((stripped, usize::MAX));
        }
        for (needle, skip) in tries {
            let Some((s, e)) = self.find(&needle, self.cursor) else { continue };
            if std::env::var_os("OXI_DOCLING_DEBUG").is_some() {
                let none = (s..e).filter(|&k| self.pt.locs[k].is_none()).count();
                eprintln!("[match] {needle:?} at {s}..{e} cursor={} missing_locs={none}", self.cursor);
            }
            self.cursor = e;
            let len = orig.chars().count();
            let mut by_page: BTreeMap<usize, (Rect, usize, usize)> = BTreeMap::new();
            for (j, idx) in (s..e).enumerate() {
                let Some((page, r)) = self.pt.locs[idx] else { continue };
                let opos = if skip == usize::MAX { None } else { kept.get(skip + j).map(|(p, _)| *p) };
                let entry = by_page.entry(page).or_insert((r, usize::MAX, 0));
                entry.0 = entry.0.union(r);
                if let Some(p) = opos {
                    entry.1 = entry.1.min(p);
                    entry.2 = entry.2.max(p + 1);
                }
            }
            if let Some(ms) = self.pt.markers.get(&s) {
                for &(page, r) in ms {
                    by_page.entry(page).and_modify(|e| e.0 = e.0.union(r)).or_insert((r, usize::MAX, 0));
                }
            }
            if by_page.is_empty() {
                return None;
            }
            let many = by_page.len() > 1;
            return Some(by_page.into_iter().map(|(page, (r, a, b))| {
                let span = if many && a <= b && b > 0 { (a, b) } else { (0, len) };
                prov(page, r, span)
            }).collect());
        }
        None
    }
}

/// Union of children's prov per page (for items without text of their own).
fn union_of(provs: &[Value]) -> Vec<Value> {
    let mut m: BTreeMap<usize, Rect> = BTreeMap::new();
    for p in provs {
        let page = p["page_no"].as_u64().unwrap() as usize;
        let b = &p["bbox"];
        let r = Rect { l: b["l"].as_f64().unwrap() as f32, t: b["t"].as_f64().unwrap() as f32,
                       r: b["r"].as_f64().unwrap() as f32, b: b["b"].as_f64().unwrap() as f32 };
        m.entry(page).and_modify(|u| *u = u.union(r)).or_insert(r);
    }
    m.into_iter().map(|(p, r)| prov(p, r, (0, 0))).collect()
}

fn resolve(d: &Value, r: &str) -> (&'static str, usize) {
    let _ = d;
    let mut it = r.trim_start_matches("#/").split('/');
    let kind = match it.next() {
        Some("texts") => "texts",
        Some("tables") => "tables",
        Some("pictures") => "pictures",
        Some("groups") => "groups",
        Some("key_value_items") => "key_value_items",
        Some("form_items") => "form_items",
        _ => "",
    };
    (kind, it.next().and_then(|i| i.parse().ok()).unwrap_or(usize::MAX))
}

struct Walk<'a, 'm> {
    m: &'m mut Matcher<'a>,
    furniture: &'a BTreeMap<(usize, bool), Vec<(String, Rect)>>,
    furniture_pictures: &'a [(usize, Rect)],
    provs: HashMap<(String, usize), Vec<Value>>,
    cell_boxes: HashMap<usize, Vec<Option<Value>>>,
}

impl Walk<'_, '_> {
    fn node(&mut self, d: &Value, r: &str) {
        let (kind, i) = resolve(d, r);
        if kind.is_empty() {
            return;
        }
        let item = &d[kind][i];
        match kind {
            "texts" => {
                let label = item["label"].as_str().unwrap_or("");
                let orig = item["orig"].as_str().or(item["text"].as_str()).unwrap_or("");
                let layer = item["content_layer"].as_str().unwrap_or("body");
                let found = if label == "formula" {
                    let at = self.m.formula_i;
                    self.m.formula_i += 1;
                    let len = item["text"].as_str().unwrap_or("").chars().count();
                    self.m.pt.formulas.get(at).filter(|l| !l.is_empty())
                        .map(|l| union_by_page(l.iter().copied()).into_iter().map(|(p, r)| prov(p, r, (0, len))).collect())
                } else if layer == "furniture" {
                    furniture_prov(self.furniture, orig)
                } else if orig.chars().any(|c| !ignorable(c)) {
                    self.m.text(orig)
                } else {
                    None
                };
                for c in item["children"].as_array().into_iter().flatten() {
                    self.node(d, c["$ref"].as_str().unwrap_or(""));
                }
                let found = found.or_else(|| {
                    // No text of its own: the union of its children.
                    let mut acc = vec![];
                    collect_child_provs(d, item, &self.provs, &mut acc);
                    (!acc.is_empty()).then(|| union_of(&acc))
                }).or_else(|| {
                    // An empty paragraph: the empty line Oxi laid out for it.
                    let blank = !orig.chars().any(|c| !ignorable(c)) && label != "formula" && layer == "body";
                    if blank { self.m.empty() } else { None }
                });
                if let Some(p) = found {
                    self.provs.insert((kind.into(), i), p);
                }
            }
            "pictures" if item["content_layer"] == "furniture" => {
                // Header/footer pictures: every page where one is painted.
                let ps: Vec<Value> = self.furniture_pictures.iter().map(|&(page, r)| prov(page, r, (0, 0))).collect();
                if !ps.is_empty() {
                    self.provs.insert((kind.into(), i), ps);
                }
            }
            "pictures" => {
                if let Some((page, r)) = self.m.picture() {
                    self.provs.insert((kind.into(), i), vec![prov(page, r, (0, 0))]);
                }
                for c in item["children"].as_array().into_iter().flatten() {
                    self.node(d, c["$ref"].as_str().unwrap_or(""));
                }
            }
            "tables" => {
                // Cells matched in order from the cursor, without moving it: a
                // rich cell's own items are matched again right after.
                let saved = self.m.cursor;
                let mut all = vec![];
                let mut boxes = vec![];
                for c in item["data"]["table_cells"].as_array().into_iter().flatten() {
                    let t = c["text"].as_str().unwrap_or("");
                    let p = if t.chars().any(|c| !ignorable(c)) { self.m.text(t) } else { None };
                    boxes.push(p.as_ref().and_then(|v| v.first()).map(|f| f["bbox"].clone()));
                    all.extend(p.unwrap_or_default());
                }
                let end = self.m.cursor;
                self.m.cursor = saved;
                let has_children = item["children"].as_array().is_some_and(|c| !c.is_empty());
                for c in item["children"].as_array().into_iter().flatten() {
                    self.node(d, c["$ref"].as_str().unwrap_or(""));
                }
                if !has_children {
                    self.m.cursor = end;
                } else {
                    self.m.cursor = self.m.cursor.max(end);
                }
                if !all.is_empty() {
                    self.provs.insert((kind.into(), i), union_of(&all));
                }
                self.cell_boxes.insert(i, boxes);
            }
            _ => {
                for c in item["children"].as_array().into_iter().flatten() {
                    self.node(d, c["$ref"].as_str().unwrap_or(""));
                }
            }
        }
    }
}

fn collect_child_provs(d: &Value, item: &Value, provs: &HashMap<(String, usize), Vec<Value>>, acc: &mut Vec<Value>) {
    for c in item["children"].as_array().into_iter().flatten() {
        let (kind, i) = resolve(d, c["$ref"].as_str().unwrap_or(""));
        if kind.is_empty() {
            continue;
        }
        if let Some(p) = provs.get(&(kind.to_string(), i)) {
            acc.extend(p.iter().cloned());
        }
        collect_child_provs(d, &d[kind][i], provs, acc);
    }
}

/// A header/footer item: every page whose header/footer fragments contain its text.
fn furniture_prov(furn: &BTreeMap<(usize, bool), Vec<(String, Rect)>>, orig: &str) -> Option<Vec<Value>> {
    let want: String = orig.chars().filter(|c| !ignorable(*c)).collect();
    if want.is_empty() {
        return None;
    }
    let len = orig.chars().count();
    let mut out = vec![];
    for (&(page, _), frs) in furn {
        // Concatenate fragments, remembering which fragment each char came from.
        let mut s = String::new();
        let mut owner = vec![];
        for (k, (t, _)) in frs.iter().enumerate() {
            for c in t.chars().filter(|c| !ignorable(*c)) {
                s.push(c);
                owner.push(k);
            }
        }
        if let Some(b) = s.find(&want) {
            let start = s[..b].chars().count();
            let n = want.chars().count();
            let r = owner[start..start + n].iter().map(|&k| frs[k].1).reduce(Rect::union).unwrap();
            out.push(prov(page, r, (0, len)));
        }
    }
    (!out.is_empty()).then_some(out)
}

/// A list number or bullet the layout painted apart from its paragraph's
/// first fragment (a text-box list, a first line that grew taller): it joins
/// the paragraph starting on its line to the right, else just below it.
fn attach_loose_markers(pt: &mut Painted, ix: &mut Index) {
    let candidates: Vec<(usize, usize)> = ix.events.iter().flat_map(|(&page, evs)| evs.iter().filter_map(move |e| match e {
        crate::Ev::Free(i) => Some((page, *i)),
        _ => None,
    })).collect();
    let mut loose: Vec<(usize, Rect)> = vec![];
    for (page, i) in candidates {
        let f = &mut ix.free[i];
        if !f.used && !f.has_run && f.text.trim().chars().count() <= 12 {
            f.used = true;
            loose.push((page, f.rect));
        }
    }
    for (page, m) in loose {
        let mid = (m.t + m.b) / 2.0;
        let h = m.b - m.t;
        let same_line = pt.para_starts.iter().filter(|(_, (p, r))| *p == page && r.l >= m.l - 0.5 && r.t <= mid && mid <= r.b)
            .min_by(|a, b| a.1 .1.l.total_cmp(&b.1 .1.l));
        let below = pt.para_starts.iter().filter(|(_, (p, r))| *p == page && r.l >= m.l - 0.5 && r.t >= m.t - 0.5 && r.t - m.b < 2.0 * h)
            .min_by(|a, b| a.1 .1.t.total_cmp(&b.1 .1.t));
        if let Some(&(start, _)) = same_line.or(below) {
            pt.markers.entry(start).or_default().push((page, m));
        }
    }
}

/// Fill `prov`, `pages` and table cell boxes of docling's own DoclingDocument
/// for `data` (the same .docx) from Oxi's layout.
pub fn annotate(data: &[u8], mut d: Value) -> Result<(Value, AnnotateStats), String> {
    let doc = oxidocs_core::parse_docx(data).map_err(|e| format!("parse: {e:?}"))?;
    let res: LayoutResult = LayoutEngine::for_document(&doc).with_show_revisions(ShowRevisions::Final).layout(&doc);
    let mut ix = index(&doc, &res);
    let mut pt = build(&doc, &mut ix);
    ix.furniture_leftovers();
    attach_loose_markers(&mut pt, &mut ix);
    let mut m = Matcher::new(&pt);
    let mut w = Walk { m: &mut m, furniture: &ix.furniture, furniture_pictures: &ix.furniture_pictures, provs: HashMap::new(), cell_boxes: HashMap::new() };
    for root in ["body", "furniture"] {
        let kids: Vec<String> = d[root]["children"].as_array().into_iter().flatten()
            .filter_map(|c| c["$ref"].as_str().map(String::from)).collect();
        for r in kids {
            w.node(&d, &r);
        }
    }
    let (mut provs, cell_boxes) = (w.provs, w.cell_boxes);

    // A comment (notes layer) is not painted in Word's final view: it takes the
    // place of the text it annotates (the items whose `comments` point at it).
    let mut annotated: HashMap<String, Vec<Value>> = HashMap::new();
    for kind in ["texts", "tables", "pictures"] {
        for (i, it) in d[kind].as_array().into_iter().flatten().enumerate() {
            for c in it["comments"].as_array().into_iter().flatten() {
                if let (Some(r), Some(p)) = (c["$ref"].as_str(), provs.get(&(kind.to_string(), i))) {
                    annotated.entry(r.to_string()).or_default().extend(p.iter().cloned());
                }
            }
        }
    }
    for (i, it) in d["texts"].as_array().into_iter().flatten().enumerate() {
        if it["content_layer"] != "notes" || provs.contains_key(&("texts".into(), i)) {
            continue;
        }
        let own = format!("#/texts/{i}");
        let parent = it["parent"]["$ref"].as_str().unwrap_or("").to_string();
        let len = it["orig"].as_str().or(it["text"].as_str()).unwrap_or("").chars().count();
        if let Some(p) = annotated.get(&own).or_else(|| annotated.get(&parent)) {
            let u: Vec<Value> = union_of(p).into_iter().map(|mut v| { v["charspan"] = json!([0, len]); v }).collect();
            provs.insert(("texts".into(), i), u);
        }
    }

    let mut stats = AnnotateStats { pages: res.pages.len(), ..Default::default() };
    for kind in ["texts", "tables", "pictures"] {
        let n = d[kind].as_array().map_or(0, |a| a.len());
        for i in 0..n {
            let item = &mut d[kind][i];
            stats.items += 1;
            match provs.get(&(kind.to_string(), i)) {
                Some(p) => {
                    item["prov"] = json!(p);
                    stats.items_with_prov += 1;
                }
                None => {
                    let t = item["text"].as_str().unwrap_or("").to_string();
                    stats.unmatched.push((format!("#/{kind}/{i}"), t.chars().take(60).collect()));
                }
            }
            if kind == "tables" {
                if let Some(boxes) = cell_boxes.get(&i) {
                    for (c, b) in item["data"]["table_cells"].as_array_mut().into_iter().flatten().zip(boxes) {
                        if let Some(b) = b {
                            c["bbox"] = b.clone();
                        }
                    }
                }
            }
        }
    }
    let mut pages = serde_json::Map::new();
    for (i, pg) in res.pages.iter().enumerate() {
        pages.insert((i + 1).to_string(), json!({"size": {"width": pg.width, "height": pg.height}, "page_no": i + 1}));
    }
    d["pages"] = Value::Object(pages);
    Ok((d, stats))
}
