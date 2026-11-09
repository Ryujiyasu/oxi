//! Map what the renderer draws onto positions in the document model.
//!
//! The editor draws a page from oxidocs-core's reading of the file and edits
//! the oxidocs-model reading of it. A click lands on a core paragraph (a
//! section's block, or a block inside a table cell) at a run and a
//! character; an edit needs a model position (a paragraph in document order
//! and a Word character position). The two readings number paragraphs and
//! characters differently (core drops field codes, writes a line break as
//! "\n", shows ruby as its base text, may split a picture into a block of
//! its own), so both are matched by their text: paragraphs by a longest
//! common subsequence of whole texts, characters within a paragraph by one
//! of characters.

use oxidocs_core::ir::{Block, Document};
use oxidocs_model::edit::{paragraph_texts, shown_positions, Story};
use oxidocs_model::DocxModel;

/// Where a core paragraph sits, as the layout names it: the section (page)
/// index (`source_container_index`) and the block index in it
/// (`paragraph_index`; for a cell paragraph, the table's); for a paragraph
/// inside table cells, one (row, cell, block) step per enclosing nested
/// table (`cell_ancestor_path`, the block being the nested table's index in
/// its cell) and a last (row, cell, paragraph) step (`cell_row_index`,
/// `cell_col_index`, `cell_paragraph_index`, which counts paragraphs only).
#[derive(Debug, Clone, PartialEq, Eq, serde::Serialize, serde::Deserialize)]
pub struct CoreKey {
    pub section: usize,
    pub block: usize,
    #[serde(default)]
    pub cells: Vec<(usize, usize, usize)>,
}

/// One core paragraph and its text, one entry per character.
struct CorePara {
    key: CoreKey,
    /// Characters of each run, in order.
    runs: Vec<Vec<char>>,
}

impl CorePara {
    fn chars(&self) -> Vec<char> {
        self.runs.iter().flatten().copied().collect()
    }
}

fn core_paragraphs(doc: &Document) -> Vec<CorePara> {
    fn runs(p: &oxidocs_core::ir::Paragraph) -> Vec<Vec<char>> {
        p.runs.iter().map(|r| r.text.chars().collect()).collect()
    }
    /// The cells of a table at `ancestors`, whose block in the section is `block`.
    fn table(t: &oxidocs_core::ir::Table, section: usize, block: usize, ancestors: &mut Vec<(usize, usize, usize)>, out: &mut Vec<CorePara>) {
        for (r, row) in t.rows.iter().enumerate() {
            for (c, cell) in row.cells.iter().enumerate() {
                let mut para = 0;
                for (i, b) in cell.blocks.iter().enumerate() {
                    match b {
                        Block::Paragraph(p) => {
                            let mut cells = ancestors.clone();
                            cells.push((r, c, para));
                            out.push(CorePara { key: CoreKey { section, block, cells }, runs: runs(p) });
                            para += 1;
                        }
                        Block::Table(inner) => {
                            ancestors.push((r, c, i));
                            table(inner, section, block, ancestors, out);
                            ancestors.pop();
                        }
                        _ => {}
                    }
                }
            }
        }
    }
    let mut out = Vec::new();
    for (s, page) in doc.pages.iter().enumerate() {
        for (i, b) in page.blocks.iter().enumerate() {
            match b {
                Block::Paragraph(p) => out.push(CorePara { key: CoreKey { section: s, block: i, cells: vec![] }, runs: runs(p) }),
                Block::Table(t) => table(t, s, i, &mut Vec::new(), &mut out),
                _ => {}
            }
        }
    }
    out
}

/// Characters the two readings write differently but mean the same.
fn same(a: char, b: char) -> bool {
    a == b || matches!((a, b), ('\n', '\u{b}') | ('\u{b}', '\n') | ('\n', '\u{c}') | ('\n', '\u{e}'))
}

/// Pairs (i, j) of a longest common subsequence of `a` and `b` under `eq`.
fn lcs_pairs<T>(a: &[T], b: &[T], eq: impl Fn(&T, &T) -> bool) -> Vec<(usize, usize)> {
    let (n, m) = (a.len(), b.len());
    // Long paragraphs or documents: keep the table small by matching the
    // common head and tail first.
    let mut head = 0;
    while head < n && head < m && eq(&a[head], &b[head]) {
        head += 1;
    }
    let mut tail = 0;
    while tail < n - head && tail < m - head && eq(&a[n - 1 - tail], &b[m - 1 - tail]) {
        tail += 1;
    }
    let (a2, b2) = (&a[head..n - tail], &b[head..m - tail]);
    let mut pairs: Vec<(usize, usize)> = (0..head).map(|i| (i, i)).collect();
    if !a2.is_empty() && !b2.is_empty() && a2.len().saturating_mul(b2.len()) <= 25_000_000 {
        let (n2, m2) = (a2.len(), b2.len());
        let mut t = vec![0u32; (n2 + 1) * (m2 + 1)];
        let at = |i: usize, j: usize| i * (m2 + 1) + j;
        for i in (0..n2).rev() {
            for j in (0..m2).rev() {
                t[at(i, j)] = if eq(&a2[i], &b2[j]) { t[at(i + 1, j + 1)] + 1 } else { t[at(i + 1, j)].max(t[at(i, j + 1)]) };
            }
        }
        let (mut i, mut j) = (0, 0);
        while i < n2 && j < m2 {
            if eq(&a2[i], &b2[j]) {
                pairs.push((head + i, head + j));
                i += 1;
                j += 1;
            } else if t[at(i + 1, j)] >= t[at(i, j + 1)] {
                i += 1;
            } else {
                j += 1;
            }
        }
    }
    pairs.extend((0..tail).map(|k| (n - tail + k, m - tail + k)));
    pairs
}

/// Pair core paragraphs with model paragraphs: by equal text, then, between
/// two matched pairs, gaps of equal length position by position.
fn pair_paragraphs(core: &[Vec<char>], model: &[Vec<char>]) -> Vec<Option<usize>> {
    let eq = |a: &Vec<char>, b: &Vec<char>| a.len() == b.len() && a.iter().zip(b).all(|(x, y)| same(*x, *y));
    let pairs = lcs_pairs(core, model, eq);
    let mut out = vec![None; core.len()];
    let mut prev = (0usize, 0usize);
    for &(i, j) in pairs.iter().chain(std::iter::once(&(core.len(), model.len()))) {
        let (gi, gj) = (i - prev.0, j - prev.1);
        if gi == gj {
            for k in 0..gi {
                out[prev.0 + k] = Some(prev.1 + k);
            }
        }
        if i < core.len() {
            out[i] = Some(j);
        }
        prev = (i + 1, j + 1);
    }
    out
}

/// A core paragraph's place in the model: the model paragraph, and for
/// each character boundary of the core text (0..=len) the Word position.
#[derive(Debug, Clone)]
pub struct ParaMap {
    pub para: usize,
    pub offsets: Vec<usize>,
}

fn map_chars(core: &[char], shown: &[(char, usize)], end: usize) -> Vec<usize> {
    let model: Vec<char> = shown.iter().map(|(c, _)| *c).collect();
    let pairs = lcs_pairs(core, &model, |a, b| same(*a, *b));
    let mut at: Vec<Option<usize>> = vec![None; core.len() + 1];
    for &(i, j) in &pairs {
        at[i] = Some(shown[j].1);
    }
    at[core.len()] = Some(end);
    // A boundary with no matched character after it takes the position
    // just past the previous one.
    let mut out = Vec::with_capacity(core.len() + 1);
    let mut last = shown.first().map(|s| s.1).unwrap_or(0);
    for (i, a) in at.iter().enumerate() {
        let p = match a {
            Some(p) => *p,
            None if i == 0 => last,
            None => last + 1,
        };
        let p = p.min(end);
        out.push(p);
        last = p;
    }
    out
}

/// The whole correspondence for a document's body.
pub struct DocMap {
    keys: Vec<CoreKey>,
    /// Characters in each run of each core paragraph.
    run_lens: Vec<Vec<usize>>,
    maps: Vec<Option<ParaMap>>,
    /// For each model paragraph, the core paragraph mapped to it.
    reverse: Vec<Option<usize>>,
}

impl DocMap {
    pub fn new(core: &Document, model: &DocxModel) -> DocMap {
        let core_paras = core_paragraphs(core);
        let texts = paragraph_texts(model, &Story::Body).unwrap_or_default();
        let shown: Vec<Vec<(char, usize)>> = texts.iter().map(|t| shown_positions(t)).collect();
        let shown_chars: Vec<Vec<char>> = shown.iter().map(|s| s.iter().map(|(c, _)| *c).collect()).collect();
        let core_chars: Vec<Vec<char>> = core_paras.iter().map(|p| p.chars()).collect();
        let pairing = pair_paragraphs(&core_chars, &shown_chars);
        let mut reverse = vec![None; texts.len()];
        let maps = pairing
            .iter()
            .enumerate()
            .map(|(i, m)| {
                m.map(|j| {
                    reverse[j] = Some(i);
                    // The paragraph's end: after its last shown character,
                    // before any content-control end.
                    let end = match shown[j].last() {
                        Some((c, p)) => p + c.len_utf16(),
                        None => leading(&texts[j]),
                    };
                    ParaMap { para: j, offsets: map_chars(&core_chars[i], &shown[j], end) }
                })
            })
            .collect();
        let run_lens = core_paras.iter().map(|p| p.runs.iter().map(|r| r.len()).collect()).collect();
        DocMap { keys: core_paras.into_iter().map(|p| p.key).collect(), run_lens, maps, reverse }
    }

    /// The model position of a place the renderer drew: a core paragraph,
    /// a run in it and a character (counted in chars) in that run.
    pub fn to_model(&self, key: &CoreKey, run: usize, char_offset: usize) -> Option<(usize, usize)> {
        let i = self.keys.iter().position(|k| k == key)?;
        let m = self.maps[i].as_ref()?;
        let before: usize = self.run_lens[i].iter().take(run).sum();
        let k = (before + char_offset).min(m.offsets.len() - 1);
        Some((m.para, m.offsets[k]))
    }

    /// The core paragraph and character (in the paragraph's concatenated
    /// run text) nearest a model position.
    pub fn to_core(&self, para: usize, offset: usize) -> Option<(CoreKey, usize, usize)> {
        let i = (*self.reverse.get(para)?)?;
        let m = self.maps[i].as_ref()?;
        let mut k = m.offsets.iter().position(|&p| p >= offset).unwrap_or(m.offsets.len() - 1);
        // As (run, character in run): the run the character starts, or the
        // end of the last run.
        for (r, &len) in self.run_lens[i].iter().enumerate() {
            if k < len || r + 1 == self.run_lens[i].len() {
                return Some((self.keys[i].clone(), r, k.min(len)));
            }
            k -= len;
        }
        Some((self.keys[i].clone(), 0, 0))
    }

    /// How many core paragraphs found a model paragraph, of how many.
    #[cfg(test)]
    pub fn coverage(&self) -> (usize, usize) {
        (self.maps.iter().filter(|m| m.is_some()).count(), self.maps.len())
    }
}

/// Positions before a paragraph's first shown character that still belong
/// to it (content-control starts), so an empty paragraph's end is after them.
fn leading(text: &str) -> usize {
    text.chars().take_while(|c| *c == oxidocs_model::edit::CONTROL_START).count()
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn characters_map_through_field_codes_and_breaks() {
        let shown = shown_positions("a\u{13} PAGE \u{14}1\u{15}b\u{b}c");
        let core: Vec<char> = "a1b\nc".chars().collect();
        let end = "a\u{13} PAGE \u{14}1\u{15}b\u{b}c".encode_utf16().count();
        // a=0, 1 at 9, b at 11, break at 12, c at 13; end 14.
        assert_eq!(map_chars(&core, &shown, end), vec![0, 9, 11, 12, 13, 14]);
    }

    #[test]
    fn a_written_document_maps_both_ways() {
        use oxidocs_model::content::Hyperlink;
        use oxidocs_model::{Block as MBlock, Cell, CellItem, Inline, Paragraph, Props, Row, RowItem, Run, RunItem, Table};
        let run = |t: &str| Inline::Run(Run { props: None, content: vec![RunItem::Text(t.into())] });
        let para = |c: Vec<Inline>| MBlock::Paragraph(Paragraph { props: None, content: c, attrs: vec![] });
        let link = Inline::Hyperlink(Hyperlink {
            target: Some("http://example.com/".into()),
            attrs: vec![("w:history".into(), "1".into())],
            content: vec![run("link")],
        });
        let cell = CellItem::Cell(Cell { props: None, content: vec![para(vec![run("cell")])], attrs: vec![] });
        let row = RowItem::Row(Row { props_ex: None, props: None, cells: vec![cell], attrs: vec![], props_after_cells: false });
        let mut tbl = Props::new("w:tblPr");
        tbl.set(oxidocs_model::xml::Element::new("w:tblW").with_attr("w:w", "0").with_attr("w:type", "auto")).unwrap();
        let model = DocxModel {
            body: vec![
                para(vec![run("ab "), link, run(" cd")]),
                MBlock::Table(Table { props: tbl, grid: None, rows: vec![row] }),
                para(vec![run("end")]),
            ],
            ..Default::default()
        };
        let bytes = model.write(None).unwrap();
        let core = oxidocs_core::parse_docx(&bytes).unwrap();
        let model = DocxModel::read(&bytes).unwrap();
        let map = DocMap::new(&core, &model);
        assert_eq!(map.coverage().0, map.coverage().1, "every core paragraph found its model paragraph");
        // "ab link cd": the l of "link" is the 4th shown character; in
        // Word positions it follows "ab ", the field begin, the code and
        // the separator: 3 + 1 + 32 + 1 = 37.
        let first = CoreKey { section: 0, block: 0, cells: vec![] };
        let total: usize = (0..map.run_lens[0].len()).map(|r| map.run_lens[0][r]).sum();
        assert_eq!(total, 10);
        let (para, off) = map.to_model(&first, 0, 3).unwrap();
        assert_eq!((para, off), (0, 37));
        // The cell paragraph, and back again.
        let i = map.keys.iter().position(|k| !k.cells.is_empty()).unwrap();
        let (para, off) = map.to_model(&map.keys[i].clone(), 0, 2).unwrap();
        assert_eq!((para, off), (1, 2));
        let (key, run, ch) = map.to_core(1, 2).unwrap();
        assert_eq!((key.cells.is_empty(), run, ch), (false, 0, 2));
        assert_eq!(map.to_model(&CoreKey { section: 0, block: 2, cells: vec![] }, 0, 3), Some((2, 3)));
    }

    #[test]
    fn paragraphs_pair_by_text_and_by_place_in_gaps() {
        let p = |s: &str| s.chars().collect::<Vec<char>>();
        let core = vec![p("one"), p("two!"), p("three")];
        let model = vec![p("one"), p("two"), p("three"), p("")];
        assert_eq!(pair_paragraphs(&core, &model), vec![Some(0), Some(1), Some(2)]);
    }
}
