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

/// A body of "abcd" / "next" on an A4 page with 30 mm side margins
/// (text width 8504 twips), as Word's own new documents have.
fn a4_doc() -> DocxModel {
    let mut sect = Element::new("w:sectPr");
    sect.children.push(oxidocs_model::xml::Node::Element(Element::new("w:pgSz").with_attr("w:w", "11906").with_attr("w:h", "16838")));
    sect.children.push(oxidocs_model::xml::Node::Element(Element::new("w:pgMar").with_attr("w:left", "1701").with_attr("w:right", "1701")));
    DocxModel {
        body: vec![para(vec![run("abcd", false)]), para(vec![run("next", false)])],
        section: Some(Props::read("w:body", &sect).unwrap()),
        ..Default::default()
    }
}

fn texts(m: &DocxModel) -> Vec<String> {
    edit::paragraphs(m, &Story::Body).unwrap().into_iter().map(edit::paragraph_text).collect()
}

#[test]
fn insert_table_goes_where_words_does() {
    // Measured over COM (Tables.Add with Word 9 table behaviour): inside a
    // paragraph or at its end, the paragraph splits and the table goes
    // between; at its start, the table goes before it.
    let cases = [
        (0, vec!["", "", "", "", "", "", "abcd", "next"], 0),
        (2, vec!["ab", "", "", "", "", "", "", "cd", "next"], 1),
        (4, vec!["abcd", "", "", "", "", "", "", "", "next"], 1),
    ];
    for (offset, want, first_cell) in cases {
        let mut m = a4_doc();
        let (edits, cell) = edit::table_edits(&m, &Pos::body(0, offset), 2, 3).unwrap();
        apply_all(&mut m, &edits).unwrap();
        assert_eq!(texts(&m), want, "cursor at {offset}");
        assert_eq!(cell, first_cell);
    }
}

#[test]
fn insert_table_shares_the_text_width_as_word_does() {
    let mut m = a4_doc();
    let (edits, _) = edit::table_edits(&m, &Pos::body(0, 0), 1, 3).unwrap();
    apply_all(&mut m, &edits).unwrap();
    let Block::Table(t) = &m.body[0] else { panic!() };
    let cols: Vec<&str> = t.grid.as_ref().unwrap().elements().map(|g| g.attr("w:w").unwrap()).collect();
    // 8504 - 10 = 8494 = 2831 + 2831 + 2832 (Word's own numbers).
    assert_eq!(cols, ["2831", "2831", "2832"]);
    let style = &m.styles.as_ref().unwrap().styles[0];
    assert!(style.meta.iter().any(|e| e.attr("w:val") == Some("Table Grid")));
    assert_eq!(t.props.get("w:tblStyle").unwrap().attr("w:val"), Some("TableGrid"));
    // A second table reuses the style.
    let (again, _) = edit::table_edits(&m, &Pos::body(3, 0), 1, 1).unwrap();
    assert!(!again.iter().any(|e| matches!(e, Edit::DefineStyle { .. })));
}

#[test]
fn insert_link_wraps_the_selection_as_words_does() {
    let mut m = DocxModel { body: vec![para(vec![run("see abc here", false)])], ..Default::default() };
    let (edits, end) = edit::link_edits(&m, &Pos::body(0, 4), &Pos::body(0, 7), "https://example.com/").unwrap();
    apply_all(&mut m, &edits).unwrap();
    let Block::Paragraph(p) = &m.body[0] else { panic!() };
    let Inline::Hyperlink(h) = &p.content[1] else { panic!("{:?}", p.content) };
    assert_eq!(h.target.as_deref(), Some("https://example.com/"));
    assert_eq!(h.attrs, vec![("w:history".to_string(), "1".to_string())]);
    let Inline::Run(r) = &h.content[0] else { panic!() };
    assert_eq!(r.props.as_ref().unwrap().get("w:rStyle").unwrap().attr("w:val"), Some("Hyperlink"));
    // The cursor goes after the link: typing there lands outside it, before " here".
    apply(&mut m, &Edit::InsertText { at: end, text: "!".into() }).unwrap();
    let shown: String = edit::shown_positions(&edit::paragraph_texts(&m, &Story::Body).unwrap()[0]).into_iter().map(|(c, _)| c).collect();
    assert_eq!(shown, "see abc! here");
    let style = &m.styles.as_ref().unwrap().styles[0];
    assert!(style.run.as_ref().unwrap().get("w:u").is_some());
}

#[test]
fn insert_link_with_nothing_selected_types_the_address() {
    let mut m = DocxModel { body: vec![para(vec![run("go ", false)])], ..Default::default() };
    let (edits, end) = edit::link_edits(&m, &Pos::body(0, 3), &Pos::body(0, 3), "https://oxi.example/").unwrap();
    apply_all(&mut m, &edits).unwrap();
    let text = &edit::paragraph_texts(&m, &Story::Body).unwrap()[0];
    let shown: String = edit::shown_positions(text).into_iter().map(|(c, _)| c).collect();
    assert_eq!(shown, "go https://oxi.example/");
    assert_eq!(end.offset, text.encode_utf16().count());
}

#[test]
fn typing_at_either_end_of_a_links_text_stays_outside_it() {
    let m0 = DocxModel { body: vec![para(vec![run("a", false), link(vec![run("bc", false)])])], ..Default::default() };
    let text = edit::paragraph_texts(&m0, &Story::Body).unwrap()[0].clone();
    let start = text.find('\u{14}').unwrap() + 1; // first character of the link's text
    let end = text.find('\u{15}').unwrap(); // just after its last character, inside
    for (at, want) in [(start, "a!bc"), (end, "abc!")] {
        let mut m = m0.clone();
        apply(&mut m, &Edit::InsertText { at: Pos::body(0, at), text: "!".into() }).unwrap();
        let Block::Paragraph(p) = &m.body[0] else { panic!() };
        let link_text: String = p.content.iter().filter_map(|i| if let Inline::Hyperlink(h) = i { Some(h) } else { None })
            .flat_map(|h| h.content.iter()).filter_map(|i| if let Inline::Run(r) = i { Some(r) } else { None })
            .flat_map(|r| r.content.iter()).filter_map(|c| if let RunItem::Text(t) = c { Some(t.as_str()) } else { None }).collect();
        assert_eq!(link_text, "bc", "typed at {at}");
        let shown: String = edit::shown_positions(&edit::paragraph_texts(&m, &Story::Body).unwrap()[0]).into_iter().map(|(c, _)| c).collect();
        assert_eq!(shown, want);
    }
}

#[test]
fn insert_footnote_as_words_does() {
    let mut m = DocxModel { body: vec![para(vec![run("abc def", false)])], ..Default::default() };
    let (edits, end) = edit::footnote_edits(&m, &Pos::body(0, 3), "note").unwrap();
    apply_all(&mut m, &edits).unwrap();
    assert_eq!(end.offset, 4);
    // Separator, continuation separator, then the note.
    let notes = m.footnotes.as_ref().unwrap();
    assert_eq!(notes.len(), 3);
    assert_eq!(notes[2].attrs, vec![("w:id".to_string(), "1".to_string())]);
    let note = &edit::paragraph_texts(&m, &Story::Footnote(2)).unwrap()[0];
    assert!(note.ends_with(" note"), "{note:?}");
    let Block::Paragraph(p) = &m.body[0] else { panic!() };
    let Inline::Run(r) = &p.content[1] else { panic!("{:?}", p.content) };
    assert!(matches!(&r.content[0], RunItem::Mark { name, .. } if name == "w:footnoteReference"));
    // A second footnote reuses the styles and separators.
    let (again, _) = edit::footnote_edits(&m, &Pos::body(0, 0), "first").unwrap();
    assert!(!again.iter().any(|e| matches!(e, Edit::DefineStyle { .. })));
    assert!(!again.iter().any(|e| matches!(e, Edit::AddStory { attrs, .. } if attrs.iter().any(|(k, _)| k == "w:type"))));
    apply_all(&mut m, &again).unwrap();
    assert_eq!(m.footnotes.as_ref().unwrap()[3].attrs[0].1, "2");
}

#[test]
fn a_footnote_mark_keeps_the_formatting_around_it_and_typing_beside_it_is_plain() {
    // Measured: in Arial 14 text the reference run is Arial 14 plus the
    // reference style; text typed just before or after it is not superscript.
    let mut m = DocxModel { body: vec![para(vec![run("Hello ", true), run("world", true)])], ..Default::default() };
    let (edits, end) = edit::footnote_edits(&m, &Pos::body(0, 6), "note").unwrap();
    apply_all(&mut m, &edits).unwrap();
    apply(&mut m, &Edit::InsertText { at: end, text: "!".into() }).unwrap();
    apply(&mut m, &Edit::InsertText { at: Pos::body(0, 6), text: "?".into() }).unwrap();
    let Block::Paragraph(p) = &m.body[0] else { panic!() };
    let runs: Vec<(String, bool, Option<String>)> = p.content.iter().filter_map(|i| if let Inline::Run(r) = i { Some(r) } else { None }).map(|r| {
        let t: String = r.content.iter().map(|c| match c { RunItem::Text(t) => t.clone(), RunItem::Mark { .. } => "^".into(), _ => String::new() }).collect();
        let pr = r.props.as_ref();
        (t, pr.is_some_and(|p| p.get("w:b").is_some()), pr.and_then(|p| p.get("w:rStyle")).and_then(|e| e.attr("w:val")).map(str::to_string))
    }).collect();
    assert_eq!(runs, vec![
        ("Hello ?".to_string(), true, None),
        ("^".to_string(), true, Some("FootnoteReference".to_string())),
        ("!world".to_string(), true, None),
    ]);
}

#[test]
fn insert_comment_as_words_does() {
    let mut m = DocxModel { body: vec![para(vec![run("Hello world", false)])], ..Default::default() };
    // A caret inside "Hello" comments on the word (measured).
    let (edits, _) = edit::comment_edits(&m, &Pos::body(0, 2), &Pos::body(0, 2), "point", "Oxi", "O", "2026-11-09T00:00:00Z").unwrap();
    apply_all(&mut m, &edits).unwrap();
    let (edits, _) = edit::comment_edits(&m, &Pos::body(0, 6), &Pos::body(0, 11), "check this", "Oxi", "O", "2026-11-09T00:00:00Z").unwrap();
    apply_all(&mut m, &edits).unwrap();
    let Block::Paragraph(p) = &m.body[0] else { panic!() };
    let names: Vec<String> = p.content.iter().map(|i| match i {
        Inline::Marker(mk) => format!("{}#{}", mk.name, mk.attrs[0].1),
        Inline::Run(r) => r.content.iter().map(|c| match c {
            RunItem::Text(t) => t.clone(),
            RunItem::Mark { name, attrs } => format!("{name}#{}", attrs[0].1),
            _ => String::new(),
        }).collect(),
        _ => String::new(),
    }).collect();
    assert_eq!(names, vec![
        "w:commentRangeStart#0", "Hello", "w:commentRangeEnd#0", "w:commentReference#0", " ",
        "w:commentRangeStart#1", "world", "w:commentRangeEnd#1", "w:commentReference#1",
    ]);
    let comments = m.comments.as_ref().unwrap();
    assert_eq!(comments.len(), 2);
    // The comment's own mark comes first, then its text.
    assert_eq!(edit::paragraph_texts(&m, &Story::Comment(1)).unwrap()[0], "\u{5}check this");
}

#[test]
fn header_and_footer_go_to_the_section_as_words_do() {
    let mut m = a4_doc();
    let edits = edit::header_footer_edits(&m, 0, false, "Head").unwrap();
    apply_all(&mut m, &edits).unwrap();
    let edits = edit::header_footer_edits(&m, 0, true, "Foot").unwrap();
    apply_all(&mut m, &edits).unwrap();
    assert_eq!(m.headers_footers.len(), 2);
    assert_eq!(m.headers_footers[0].kind, "w:hdr");
    let refs: Vec<(String, String)> = m.section.as_ref().unwrap().children()
        .filter(|e| e.name.ends_with("Reference"))
        .map(|e| (e.name.clone(), e.attr("r:id").unwrap().to_string())).collect();
    assert_eq!(refs, vec![("w:headerReference".into(), "#0".into()), ("w:footerReference".into(), "#1".into())]);
    // The header style's tabs sit at the middle and the end of the 8504-twip text width.
    let style = m.styles.as_ref().unwrap().styles.iter().find(|s| s.attrs.iter().any(|(_, v)| v == "Header")).unwrap();
    let tabs: Vec<&str> = style.paragraph.as_ref().unwrap().get("w:tabs").unwrap().elements().map(|t| t.attr("w:pos").unwrap()).collect();
    assert_eq!(tabs, ["4252", "8504"]);
    // A second header command replaces the text.
    let edits = edit::header_footer_edits(&m, 0, false, "New head").unwrap();
    apply_all(&mut m, &edits).unwrap();
    assert_eq!(m.headers_footers.len(), 2);
    assert_eq!(edit::paragraph_texts(&m, &Story::HeaderFooter(0)).unwrap(), vec!["New head"]);
}

#[test]
fn insert_field_as_words_does() {
    let mut m = DocxModel { body: vec![para(vec![run("Page  of", true)])], ..Default::default() };
    let (edits, end) = edit::field_edits(&m, &Pos::body(0, 5), "PAGE", "1").unwrap();
    apply_all(&mut m, &edits).unwrap();
    let text = &edit::paragraph_texts(&m, &Story::Body).unwrap()[0];
    assert_eq!(text, "Page \u{13} PAGE \\* MERGEFORMAT \u{14}1\u{15} of");
    assert_eq!(end.offset, text.find(" of").unwrap());
    let shown: String = edit::shown_positions(text).into_iter().map(|(c, _)| c).collect();
    assert_eq!(shown, "Page 1 of");
    // Every run keeps the bold around it; the result is not proofread.
    let Block::Paragraph(p) = &m.body[0] else { panic!() };
    for i in &p.content {
        let Inline::Run(r) = i else { panic!() };
        assert!(r.props.as_ref().unwrap().get("w:b").is_some());
    }
}

#[test]
fn a_fields_runs_stay_apart_from_the_text_around_it() {
    // Word writes each field character and the code in runs of their own.
    let mut m = DocxModel { body: vec![para(vec![run("Hello world", false)])], ..Default::default() };
    let (edits, _) = edit::field_edits(&m, &Pos::body(0, 6), "PAGE", "1").unwrap();
    apply_all(&mut m, &edits).unwrap();
    let Block::Paragraph(p) = &m.body[0] else { panic!() };
    let mixed = p.content.iter().any(|i| matches!(i, Inline::Run(r)
        if r.content.iter().any(|c| matches!(c, RunItem::FieldChar { .. } | RunItem::InstrText(_)))
        && r.content.iter().any(|c| matches!(c, RunItem::Text(_)))));
    assert!(!mixed, "{:?}", p.content);
}
