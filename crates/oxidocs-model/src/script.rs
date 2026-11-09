// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! The edits that make a document from nothing.
//!
//! [`script`] reads a model and writes the [`Edit`]s that build it, starting
//! from an empty document, the way a person would in an editor: set up the
//! styles and page, type each paragraph's runs, wrap links and fields
//! around what they hold, insert tables and fill their cells, add headers,
//! footers and notes. Applying the script to [`DocxModel::default`] gives
//! the document back; the `recreate` example checks that for a corpus. When
//! no edit can make something, [`script`] names it.

use crate::content::{Block, Cell, CellItem, Inline, Paragraph, Revision, Row, RowItem, Sdt, Table, Wrapper};
use crate::docparts::Numbering;
use crate::edit::{self, DocumentParts, Edit, Patch, Pos, Story, StoryKind};
use crate::package::DocxModel;

/// Something in the document no edit can make yet.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Missing(pub String);

impl std::fmt::Display for Missing {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        write!(f, "no edit makes {}", self.0)
    }
}

fn missing<T>(what: impl Into<String>) -> Result<T, Missing> {
    Err(Missing(what.into()))
}

struct Gen {
    work: DocxModel,
    edits: Vec<Edit>,
}

impl Gen {
    fn push(&mut self, e: Edit) -> Result<(), Missing> {
        edit::apply(&mut self.work, &e).map_err(|x| Missing(format!("(edit refused: {})", x.0)))?;
        self.edits.push(e);
        Ok(())
    }

    fn end(&self, story: &Story, para: usize) -> Result<usize, Missing> {
        edit::paragraph_end(&self.work, story, para).map_err(|x| Missing(format!("(no paragraph: {})", x.0)))
    }

    fn pos(&self, story: &Story, para: usize, offset: usize) -> Pos {
        Pos { story: story.clone(), para, offset }
    }

    /// Fill the blocks of one container (a story or a cell) whose
    /// paragraphs start at `first`; the container holds one empty paragraph
    /// there (in the file or implied). Returns how many paragraphs the
    /// container has afterwards.
    ///
    /// The container keeps a free empty paragraph at its end while it is
    /// built, as an editor keeps the cursor's paragraph: each paragraph is
    /// made by breaking the free one at its start and filling the first
    /// half, and tables and markers go in before it. At the end the free
    /// paragraph goes.
    fn blocks(&mut self, story: &Story, blocks: &[Block], first: usize) -> Result<usize, Missing> {
        let before = self.count(story)?;
        let mut leaves = Vec::new();
        flatten(blocks, &mut leaves);
        let mut cur = first;
        // Blocks made so far in each open container, innermost last.
        let mut made: Vec<usize> = vec![0];
        for leaf in leaves {
            match leaf {
                Leaf::Block(Block::Paragraph(p)) => {
                    self.push(Edit::SplitParagraph { at: self.pos(story, cur, 0) })?;
                    self.paragraph(story, cur, p)?;
                    cur += 1;
                }
                Leaf::Block(Block::Table(t)) => {
                    let shell = Table { props: t.props.clone(), grid: t.grid.clone(), rows: empty_rows(&t.rows) };
                    self.push(Edit::InsertTable { story: story.clone(), para: cur, after: false, table: shell })?;
                    cur += self.rows(story, &t.rows, cur)?;
                }
                Leaf::Block(b) => {
                    self.push(Edit::InsertBlock { story: story.clone(), para: cur, block: b.clone() })?;
                }
                Leaf::Open => {
                    made.push(0);
                    continue;
                }
                Leaf::Close(w) => {
                    let count = made.pop().unwrap();
                    if count == 0 {
                        return missing("an empty block container");
                    }
                    self.push(Edit::WrapBlocks { story: story.clone(), para: cur, count, wrapper: w })?;
                }
            }
            *made.last_mut().unwrap() += 1;
        }
        if !edit::implied(&self.work, story, cur).map_err(|x| Missing(x.0))? {
            self.push(Edit::RemoveParagraph { story: story.clone(), para: cur })?;
        }
        Ok(self.count(story)? + 1 - before)
    }

    fn count(&self, story: &Story) -> Result<usize, Missing> {
        Ok(edit::paragraphs(&self.work, story).map_err(|x| Missing(x.0))?.len())
    }

    fn rows(&mut self, story: &Story, rows: &[RowItem], first: usize) -> Result<usize, Missing> {
        let mut at = first;
        for r in rows {
            match r {
                RowItem::Row(row) => {
                    for c in &row.cells {
                        at += self.cell(story, c, at)?;
                    }
                }
                RowItem::Marker(_) => {}
                RowItem::Sdt(x) => at += self.rows(story, &x.content, at)?,
                RowItem::CustomXml(x) => at += self.rows(story, &x.content, at)?,
                RowItem::Revision(x) => at += self.rows(story, &x.content, at)?,
            }
        }
        Ok(at - first)
    }

    fn cell(&mut self, story: &Story, c: &CellItem, first: usize) -> Result<usize, Missing> {
        match c {
            CellItem::Cell(cell) => self.blocks(story, &cell.content, first),
            CellItem::Marker(_) => Ok(0),
            CellItem::Sdt(x) => self.cells(story, &x.content, first),
            CellItem::CustomXml(x) => self.cells(story, &x.content, first),
            CellItem::Revision(x) => self.cells(story, &x.content, first),
        }
    }

    fn cells(&mut self, story: &Story, cells: &[CellItem], first: usize) -> Result<usize, Missing> {
        let mut at = first;
        for c in cells {
            at += self.cell(story, c, at)?;
        }
        Ok(at - first)
    }

    fn paragraph(&mut self, story: &Story, para: usize, p: &Paragraph) -> Result<(), Missing> {
        let here = edit::paragraphs(&self.work, story).map_err(|x| Missing(x.0))?[para].clone();
        if here.props != p.props {
            let set = p.props.as_ref().map(|pp| pp.children().cloned().collect()).unwrap_or_default();
            let patch = Patch { set, clear: true, ..Default::default() };
            if p.props.as_ref().is_some_and(|pp| !pp.element().attrs.is_empty()) {
                return missing("attributes on paragraph properties");
            }
            let at = self.pos(story, para, 0);
            self.push(Edit::FormatParagraphs { from: at.clone(), to: at, patch })?;
        }
        if here.attrs != p.attrs {
            self.push(Edit::SetParagraphAttrs { story: story.clone(), para, attrs: p.attrs.clone() })?;
        }
        for i in &p.content {
            self.inline(story, para, i)?;
        }
        Ok(())
    }

    fn inline(&mut self, story: &Story, para: usize, i: &Inline) -> Result<(), Missing> {
        match i {
            Inline::Run(_) | Inline::Marker(_) | Inline::Math(_) | Inline::Loose(_) => {
                let end = self.end(story, para)?;
                self.push(Edit::InsertInline { at: self.pos(story, para, end), inline: i.clone() })
            }
            _ => {
                let from = self.end(story, para)?;
                // What of no width ends the paragraph so far stays before the container.
                let here = &edit::paragraphs(&self.work, story).map_err(|x| Missing(x.0))?[para].content;
                let outside = here.iter().rev().take_while(|x| edit::inline_width(x) == 0).count();
                for c in children(i) {
                    self.inline(story, para, c)?;
                }
                let to = self.end(story, para)?;
                self.push(Edit::Wrap { from: self.pos(story, para, from), to: self.pos(story, para, to), wrapper: shell(i), outside })
            }
        }
    }

    /// Build a story added empty: its paragraphs start at 0.
    fn story(&mut self, story: Story, blocks: &[Block]) -> Result<(), Missing> {
        self.blocks(&story, blocks, 0).map(|_| ())
    }
}

enum Leaf<'a> {
    Block(&'a Block),
    Open,
    /// The end of a block container; the container with no content.
    Close(Block),
}

/// Blocks in order, with block containers as an open and a close around
/// what they hold.
fn flatten<'a>(blocks: &'a [Block], out: &mut Vec<Leaf<'a>>) {
    for b in blocks {
        let (inner, shell) = match b {
            Block::Sdt(s) => (&s.content, Block::Sdt(Sdt { content: Vec::new(), ..s.clone() })),
            Block::CustomXml(w) => (&w.content, Block::CustomXml(Wrapper { content: Vec::new(), ..w.clone() })),
            Block::Revision(r) => (&r.content, Block::Revision(Revision { content: Vec::new(), ..r.clone() })),
            _ => {
                out.push(Leaf::Block(b));
                continue;
            }
        };
        out.push(Leaf::Open);
        flatten(inner, out);
        out.push(Leaf::Close(shell));
    }
}

fn children(i: &Inline) -> &[Inline] {
    match i {
        Inline::Hyperlink(h) => &h.content,
        Inline::SimpleField(f) => &f.content,
        Inline::Sdt(s) => &s.content,
        Inline::SmartTag(w) | Inline::CustomXml(w) | Inline::Bidi(w) => &w.content,
        Inline::Revision(r) => &r.content,
        _ => &[],
    }
}

/// A container with its content taken out.
fn shell(i: &Inline) -> Inline {
    let mut s = i.clone();
    match &mut s {
        Inline::Hyperlink(h) => h.content.clear(),
        Inline::SimpleField(f) => f.content.clear(),
        Inline::Sdt(x) => x.content.clear(),
        Inline::SmartTag(w) | Inline::CustomXml(w) | Inline::Bidi(w) => w.content.clear(),
        Inline::Revision(r) => r.content.clear(),
        _ => {}
    }
    s
}

/// A table's rows with every cell emptied: the table as inserted, before
/// its cells are filled.
fn empty_rows(rows: &[RowItem]) -> Vec<RowItem> {
    fn cells(c: &[CellItem]) -> Vec<CellItem> {
        c.iter()
            .map(|c| match c {
                CellItem::Cell(cell) => CellItem::Cell(Cell { content: Vec::new(), ..cell.clone() }),
                CellItem::Sdt(x) => CellItem::Sdt(Sdt { content: cells(&x.content), ..x.clone() }),
                CellItem::CustomXml(x) => CellItem::CustomXml(Wrapper { content: cells(&x.content), ..x.clone() }),
                CellItem::Revision(x) => CellItem::Revision(Revision { content: cells(&x.content), ..x.clone() }),
                CellItem::Marker(m) => CellItem::Marker(m.clone()),
            })
            .collect()
    }
    rows.iter()
        .map(|r| match r {
            RowItem::Row(row) => RowItem::Row(Row { cells: cells(&row.cells), ..row.clone() }),
            RowItem::Sdt(x) => RowItem::Sdt(Sdt { content: empty_rows(&x.content), ..x.clone() }),
            RowItem::CustomXml(x) => RowItem::CustomXml(Wrapper { content: empty_rows(&x.content), ..x.clone() }),
            RowItem::Revision(x) => RowItem::Revision(Revision { content: empty_rows(&x.content), ..x.clone() }),
            RowItem::Marker(m) => RowItem::Marker(m.clone()),
        })
        .collect()
}

fn id(attrs: &[(String, String)], name: &str) -> Option<String> {
    attrs.iter().find(|(k, _)| k == name).map(|(_, v)| v.clone())
}

/// Whether every definition has an id, and no two share one.
fn unique<T>(list: &[T], key: impl Fn(&T) -> Option<String>) -> bool {
    let mut seen = std::collections::HashSet::new();
    list.iter().all(|x| key(x).is_some_and(|k| seen.insert(k)))
}

fn document_parts(m: &DocxModel) -> DocumentParts {
    DocumentParts {
        background: m.background.clone(),
        styles: m.styles.clone(),
        styles_with_effects: m.styles_with_effects.clone(),
        numbering: m.numbering.clone(),
        settings: m.settings.clone(),
        web_settings: m.web_settings.clone(),
        font_table: m.font_table.clone(),
        people: m.people.clone(),
        theme: m.theme.clone(),
        custom_xml: m.custom_xml.clone(),
        glossary: m.glossary.clone(),
        app_data: m.app_data.clone(),
        comments_extended: m.comments_extended.clone(),
        comments_ids: m.comments_ids.clone(),
        comments_extensible: m.comments_extensible.clone(),
        footnotes: m.footnotes.is_some(),
        endnotes: m.endnotes.is_some(),
        comments: m.comments.is_some(),
    }
}

/// The edits that make `model` from [`DocxModel::default`].
pub fn script(model: &DocxModel) -> Result<Vec<Edit>, Missing> {
    let mut g = Gen { work: DocxModel::default(), edits: Vec::new() };
    let mut parts = document_parts(model);
    // Styles and lists are defined one by one, as an editor's dialogs do;
    // a document that gives two definitions one id gets its part whole.
    let mut defs = Vec::new();
    for (with_effects, st) in [(false, &model.styles), (true, &model.styles_with_effects)] {
        let Some(st) = st else { continue };
        if !unique(&st.styles, |x| id(&x.attrs, "w:styleId")) {
            continue;
        }
        if with_effects {
            parts.styles_with_effects = None;
        } else {
            parts.styles = None;
        }
        defs.push(Edit::SetStyleDefaults {
            with_effects,
            run: st.default_run.clone(),
            paragraph: st.default_paragraph.clone(),
            declared: st.has_defaults,
            latent: st.latent.clone(),
        });
        for style in &st.styles {
            defs.push(Edit::DefineStyle { with_effects, style: Box::new(style.clone()) });
        }
    }
    if let Some(n) = &model.numbering {
        if unique(&n.abstracts, |x| id(&x.attrs, "w:abstractNumId")) && unique(&n.nums, |x| id(&x.attrs, "w:numId")) {
            // Picture bullets and the like stay with the part.
            parts.numbering = Some(Numbering { abstracts: Vec::new(), nums: Vec::new(), ..n.clone() });
            for a in &n.abstracts {
                defs.push(Edit::DefineList { list: Box::new(a.clone()) });
            }
            for x in &n.nums {
                defs.push(Edit::DefineListInstance { num: x.clone() });
            }
        }
    }
    g.push(Edit::SetDocumentParts { parts: Box::new(parts) })?;
    for e in defs {
        g.push(e)?;
    }
    for (i, hf) in model.headers_footers.iter().enumerate() {
        g.push(Edit::AddStory { story: StoryKind::HeaderFooter { kind: hf.kind.clone() }, attrs: vec![] })?;
        g.story(Story::HeaderFooter(i), &hf.content)?;
    }
    let notes = [
        (&model.footnotes, StoryKind::Footnote, Story::Footnote as fn(usize) -> Story),
        (&model.endnotes, StoryKind::Endnote, Story::Endnote),
        (&model.comments, StoryKind::Comment, Story::Comment),
    ];
    for (list, kind, story) in notes {
        for (i, n) in list.iter().flatten().enumerate() {
            g.push(Edit::AddStory { story: kind.clone(), attrs: n.attrs.clone() })?;
            g.story(story(i), &n.content)?;
        }
    }
    g.story(Story::Body, &model.body)?;
    if model.section.is_some() || model.section_at.is_some() {
        g.push(Edit::SetSection { props: model.section.clone(), at: model.section_at })?;
    }
    Ok(g.edits)
}
