// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

use super::*;
use std::collections::{BTreeMap, BTreeSet, HashMap};

type StyleMap = HashMap<String, String>;

#[derive(Clone, PartialEq, Eq)]
pub(super) struct PageReferences {
    pub map: StyleMap,
    pub first_block: usize,
}

#[derive(Default)]
pub(super) struct StylePagination {
    pub referenced: BTreeSet<String>,
    pub known: BTreeMap<(usize, usize), PageReferences>,
    pub observed: BTreeMap<(usize, usize), PageReferences>,
    seen: Vec<BTreeMap<(usize, usize), PageReferences>>,
}

impl StylePagination {
    pub fn initialize(&mut self, doc: &Document) {
        fn visit(blocks: &[Block], ids: &mut BTreeSet<String>) {
            for b in blocks {
                match b {
                    Block::Paragraph(p) => for r in &p.runs {
                        if let Some(id) = r.style.styleref.as_deref().and_then(style_reference::style_id) { ids.insert(id); }
                    },
                    Block::Table(t) => for row in &t.rows { for c in &row.cells { visit(&c.blocks, ids); } },
                    _ => {}
                }
            }
        }
        for p in &doc.pages {
            for b in [&p.header, &p.header_first, &p.header_even, &p.footer, &p.footer_first, &p.footer_even] {
                visit(b, &mut self.referenced);
            }
            for r in &p.header_runs {
                for b in [&r.header, &r.header_first, &r.header_even, &r.footer, &r.footer_first, &r.footer_even] {
                    visit(b, &mut self.referenced);
                }
            }
        }
    }

    pub fn advance(&mut self) -> bool {
        if self.known == self.observed { return false; }
        if self.seen.iter().any(|m| m == &self.observed) {
            eprintln!("OXI: style-reference pagination cycle; last complete layout retained");
            return false;
        }
        self.seen.push(self.known.clone());
        self.known = self.observed.clone();
        true
    }

    pub fn map(&self, ir: usize, page: usize) -> Option<&StyleMap> {
        self.known.get(&(ir, page)).map(|p| &p.map)
    }

    pub fn observe(&mut self, ir: usize, page: &Page, laid_out: &[LayoutPage], starts: &[usize], initial: &StyleMap) {
        if self.referenced.is_empty() { return; }
        let mut first = vec![StyleMap::new(); laid_out.len()];
        let mut last = first.clone();
        let mut first_blocks = vec![usize::MAX; laid_out.len()];
        for (pno, lp) in laid_out.iter().enumerate() {
            for e in &lp.elements {
                if let Some(b) = e.paragraph_index { first_blocks[pno] = first_blocks[pno].min(b); }
            }
        }
        let mut record = |pno: usize, block: usize, id: &str, text: &str| {
            if pno >= first.len() || !self.referenced.contains(id) { return; }
            first[pno].entry(id.to_owned()).or_insert_with(|| text.to_owned());
            last[pno].insert(id.to_owned(), text.to_owned());
            first_blocks[pno] = first_blocks[pno].min(block);
        };
        for (bi, block) in page.blocks.iter().enumerate() {
            let fallback = starts.get(bi).copied().unwrap_or(0).min(laid_out.len().saturating_sub(1));
            match block {
                Block::Paragraph(p) => {
                    let elements = paragraph_elements(laid_out, bi, None);
                    locate(p, bi, fallback, &elements, &mut record);
                }
                Block::Table(t) => visit_table(t, bi, &[], fallback, laid_out, &mut record),
                _ => {}
            }
        }
        let mut previous: StyleMap = initial.iter().filter(|(id,_)| self.referenced.contains(*id)).map(|(k,v)|(k.clone(),v.clone())).collect();
        for pno in 0..laid_out.len() {
            let mut map = previous.clone();
            map.extend(first[pno].clone());
            self.observed.insert((ir, pno+1), PageReferences { map, first_block: first_blocks[pno] });
            previous.extend(last[pno].clone());
        }
    }
}

fn paragraph_elements<'a>(pages: &'a [LayoutPage], block: usize, cell: Option<&CellFlowKey>) -> Vec<(usize, &'a LayoutElement)> {
    pages.iter().enumerate().flat_map(|(pi, p)| p.elements.iter().filter_map(move |e| {
        if e.paragraph_index != Some(block) || !matches!(e.content, LayoutContent::Text { .. }) { return None; }
        let matching = match cell { Some(key) => e.cell_flow_key().as_ref() == Some(key), None => e.cell_paragraph_index.is_none() };
        matching.then_some((pi,e))
    })).collect()
}

fn visit_table(table: &Table, block: usize, path: &[(usize,usize,usize)], fallback: usize, pages: &[LayoutPage],
    record: &mut impl FnMut(usize,usize,&str,&str)) {
    for (ri,row) in table.rows.iter().enumerate() {
        for (ci,cell) in row.cells.iter().enumerate() {
            let mut para = 0;
            for (bi,b) in cell.blocks.iter().enumerate() {
                match b {
                    Block::Paragraph(p) => {
                        let key = (path.to_vec(),ri,ci,para);
                        let elements = paragraph_elements(pages,block,Some(&key));
                        locate(p,block,fallback,&elements,record);
                        para += 1;
                    }
                    Block::Table(t) => {
                        let mut next = path.to_vec();next.push((ri,ci,bi));
                        visit_table(t,block,&next,fallback,pages,record);
                    }
                    _ => {}
                }
            }
        }
    }
}

fn locate(para: &Paragraph, block: usize, fallback: usize, elements: &[(usize,&LayoutElement)],
    record: &mut impl FnMut(usize,usize,&str,&str)) {
    let para_pages: BTreeSet<usize> = elements.iter().map(|(p,_)|*p).collect();
    let start_page = para_pages.first().copied().unwrap_or(fallback);
    if let Some(id) = para.style.style_id.as_deref() {
        let text: String = para.runs.iter().map(|r|r.text.as_str()).collect();
        if para_pages.is_empty() { record(start_page,block,id,text.trim()); }
        else { for p in &para_pages { record(*p,block,id,text.trim()); } }
    }
    // The cell wrapper emits fragments without run indices. Project its
    // original text in emission order onto source runs; whitespace omitted
    // from painted text cannot move a following character to an earlier page.
    let source: Vec<(char,usize)> = para.runs.iter().enumerate().flat_map(|(i,r)|
        r.text.chars().filter(|c|!c.is_whitespace()).map(move |c|(c.to_ascii_lowercase(),i))).collect();
    let mut run_pages = vec![BTreeSet::new();para.runs.len()];
    let mut offset = 0;
    for (p,e) in elements {
        if let Some(run) = e.run_index {
            if let Some(set) = run_pages.get_mut(run) { set.insert(*p); }
            continue;
        }
        let LayoutContent::Text { text, .. } = &e.content else { continue; };
        let original = e.source_text.as_deref().unwrap_or(text);
        let chars: Vec<char> = original.chars().filter(|c|!c.is_whitespace()).map(|c|c.to_ascii_lowercase()).collect();
        if chars.is_empty() { continue; }
        let found = (offset..=source.len().saturating_sub(chars.len())).find(|&i|
            source.get(i..i+chars.len()).is_some_and(|slice|slice.iter().map(|(c,_)|*c).eq(chars.iter().copied())));
        if let Some(i) = found {
            for (_,run) in &source[i..i+chars.len()] { run_pages[*run].insert(*p); }
            offset = i+chars.len();
        }
    }
    let mut i = 0;
    while i < para.runs.len() {
        let Some(id) = para.runs[i].style.char_style_id.as_deref() else { i+=1; continue; };
        let start = i;let mut text=String::new();let mut found=BTreeSet::new();
        while i<para.runs.len() && para.runs[i].style.char_style_id.as_deref()==Some(id) {
            text.push_str(&para.runs[i].text);found.extend(run_pages[i].iter().copied());i+=1;
        }
        if text.is_empty() { continue; }
        if found.is_empty() {
            // A preserved styled space is an occurrence even without painted
            // ink. Its next visible run locates its flow position; at the end
            // of a paragraph the last preceding run supplies that position.
            let adjacent = run_pages[i..].iter().find_map(|s|s.first().copied())
                .or_else(||run_pages[..start].iter().rev().find_map(|s|s.last().copied())).unwrap_or(start_page);
            found.insert(adjacent);
        }
        for p in found { record(p,block,id,text.trim()); }
    }
}

thread_local! {
    static BOXES: std::cell::RefCell<BTreeMap<usize,(f32,f32)>> = std::cell::RefCell::new(BTreeMap::new());
}

pub(super) struct GeometryScope(BTreeMap<usize,(f32,f32)>);
impl GeometryScope {
    pub fn new(boxes: BTreeMap<usize,(f32,f32)>) -> Self { Self(BOXES.with(|b|b.replace(boxes))) }
}
impl Drop for GeometryScope {
    fn drop(&mut self) { BOXES.with(|b|b.replace(std::mem::take(&mut self.0))); }
}
pub(super) fn geometry(page: usize) -> Option<(f32,f32)> { BOXES.with(|b|b.borrow().get(&page).copied()) }
