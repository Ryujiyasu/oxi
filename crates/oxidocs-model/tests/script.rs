// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! Documents made from nothing with edits: the structural edits (leaves,
//! containers, tables, block containers) and the script that strings them
//! together.

use oxidocs_model::content::{Hyperlink, Marker, Revision, Sdt};
use oxidocs_model::edit::{self, apply, apply_all, Edit, Pos, Story};
use oxidocs_model::xml::Element;
use oxidocs_model::{script, Block, Cell, CellItem, DocxModel, Inline, Paragraph, Props, Row, RowItem, Run, RunItem, Table};

fn run(text: &str, bold: bool) -> Inline {
    let props = bold.then(|| {
        let mut p = Props::new("w:rPr");
        p.set(Element::new("w:b")).unwrap();
        p
    });
    Inline::Run(Run { props, content: vec![RunItem::Text(text.into())] })
}

fn marker(name: &str, id: &str) -> Inline {
    Inline::Marker(Marker { name: name.into(), attrs: vec![("w:id".into(), id.into())] })
}

fn para(content: Vec<Inline>) -> Block {
    Block::Paragraph(Paragraph { props: None, content, attrs: vec![] })
}

fn cell(content: Vec<Block>) -> CellItem {
    CellItem::Cell(Cell { props: None, content, attrs: vec![] })
}

fn table(cells: Vec<CellItem>) -> Block {
    let row = RowItem::Row(Row { props_ex: None, props: None, cells, attrs: vec![], props_after_cells: false });
    Block::Table(Table { props: Props::new("w:tblPr"), grid: None, rows: vec![row] })
}

fn link(content: Vec<Inline>) -> Inline {
    Inline::Hyperlink(Hyperlink { target: Some("https://example.com/".into()), attrs: vec![], content })
}

fn deletion() -> Revision<Inline> {
    Revision { name: "w:del".into(), attrs: vec![("w:id".into(), "7".into())], content: vec![] }
}

/// The script for `model`, applied to an empty document after a trip
/// through JSON.
fn remade(model: &DocxModel) -> DocxModel {
    let edits = script::script(model).unwrap();
    let json = serde_json::to_string(&edits).unwrap();
    let edits: Vec<Edit> = serde_json::from_str(&json).unwrap();
    let mut made = DocxModel::default();
    apply_all(&mut made, &edits).unwrap();
    made
}

#[test]
fn a_leaf_goes_after_what_has_no_width_at_its_position() {
    let mut m = DocxModel { body: vec![para(vec![run("ab", false), marker("w:bookmarkStart", "1")])], ..Default::default() };
    apply(&mut m, &Edit::InsertInline { at: Pos::body(0, 2), inline: run("c", true) }).unwrap();
    let Block::Paragraph(p) = &m.body[0] else { panic!() };
    assert!(matches!(&p.content[1], Inline::Marker(_)));
    assert_eq!(p.content[2], run("c", true));
}

#[test]
fn a_leaf_inside_a_hyperlinks_text_goes_into_the_link() {
    let mut m = DocxModel { body: vec![para(vec![link(vec![run("ab", false)])])], ..Default::default() };
    // The link's code takes positions 0..; its text starts after the separator.
    let start = edit::paragraph_text(edit::paragraphs(&m, &Story::Body).unwrap()[0]).find('\u{14}').unwrap() + 1;
    apply(&mut m, &Edit::InsertInline { at: Pos::body(0, start + 1), inline: marker("w:bookmarkStart", "1") }).unwrap();
    let Block::Paragraph(p) = &m.body[0] else { panic!() };
    let Inline::Hyperlink(h) = &p.content[0] else { panic!() };
    assert!(matches!(h.content[1], Inline::Marker(_)));
}

#[test]
fn wrap_leaves_the_counted_zero_width_things_outside() {
    let mut m = DocxModel { body: vec![para(vec![run("a", false), marker("w:bookmarkStart", "1"), run("b", false)])], ..Default::default() };
    let wrapper = Inline::Revision(deletion());
    apply(&mut m, &Edit::Wrap { from: Pos::body(0, 1), to: Pos::body(0, 2), wrapper: wrapper.clone(), outside: 1 }).unwrap();
    let Block::Paragraph(p) = &m.body[0] else { panic!() };
    assert!(matches!(p.content[1], Inline::Marker(_)));
    let Inline::Revision(r) = &p.content[2] else { panic!("{:?}", p.content) };
    assert_eq!(r.content, vec![run("b", false)]);
}

#[test]
fn editing_one_cell_leaves_the_other_empty_cells_as_they_were() {
    let mut m = DocxModel { body: vec![table(vec![cell(vec![]), cell(vec![])]), para(vec![])], ..Default::default() };
    apply(&mut m, &Edit::InsertText { at: Pos::body(0, 0), text: "x".into() }).unwrap();
    let Block::Table(t) = &m.body[0] else { panic!() };
    let RowItem::Row(r) = &t.rows[0] else { panic!() };
    let CellItem::Cell(second) = &r.cells[1] else { panic!() };
    assert!(second.content.is_empty());
}

#[test]
fn a_table_goes_in_before_a_paragraph_and_an_empty_paragraph_comes_out() {
    let mut m = DocxModel::default();
    let shell = table(vec![cell(vec![])]);
    let Block::Table(t) = shell.clone() else { panic!() };
    apply(&mut m, &Edit::InsertTable { story: Story::Body, para: 0, after: false, table: t }).unwrap();
    // The table's cell, then the paragraph Word gives a body ending in a table.
    assert_eq!(edit::paragraphs(&m, &Story::Body).unwrap().len(), 2);
    assert!(edit::implied(&m, &Story::Body, 1).unwrap());
    apply(&mut m, &Edit::SplitParagraph { at: Pos::body(1, 0) }).unwrap();
    apply(&mut m, &Edit::RemoveParagraph { story: Story::Body, para: 2 }).unwrap();
    assert_eq!(m.body.len(), 2);
    assert!(apply(&mut m, &Edit::InsertText { at: Pos::body(1, 0), text: "x".into() }).is_ok());
    assert!(apply(&mut m, &Edit::RemoveParagraph { story: Story::Body, para: 1 }).is_err(), "not empty");
}

#[test]
fn block_containers_wrap_the_blocks_before_a_paragraph() {
    let mut m = DocxModel { body: vec![para(vec![run("a", false)]), para(vec![run("b", false)]), para(vec![])], ..Default::default() };
    let control = Block::Sdt(Sdt { props: None, end_props: None, content: vec![] });
    apply(&mut m, &Edit::WrapBlocks { story: Story::Body, para: 2, count: 2, wrapper: control }).unwrap();
    assert_eq!(m.body.len(), 2);
    assert_eq!(edit::paragraph_texts(&m, &Story::Body).unwrap()[..2], ["\u{1c}a".to_string(), "b\u{1d}".to_string()]);
}

#[test]
fn a_script_makes_the_document_again() {
    let mut del = deletion();
    del.content = vec![run("gone", false)];
    let body = vec![
        para(vec![run("plain ", false), run("bold", true), marker("w:bookmarkStart", "0"), Inline::Revision(del)]),
        Block::Marker(Marker { name: "w:bookmarkEnd".into(), attrs: vec![("w:id".into(), "0".into())] }),
        table(vec![cell(vec![para(vec![link(vec![run("link", false)])])]), cell(vec![])]),
        Block::Sdt(Sdt { props: None, end_props: None, content: vec![para(vec![run("in a control", false)])] }),
        para(vec![]),
        table(vec![cell(vec![table(vec![cell(vec![para(vec![run("nested", true)])])])])]),
    ];
    let model = DocxModel { body, ..Default::default() };
    assert_eq!(remade(&model), model);
}

#[test]
fn what_no_edit_makes_is_named() {
    // A block container with nothing in it has no blocks to wrap.
    let body = vec![Block::Sdt(Sdt { props: None, end_props: None, content: vec![] })];
    let err = script::script(&DocxModel { body, ..Default::default() }).unwrap_err();
    assert!(err.0.contains("empty block container"), "{err}");
}
