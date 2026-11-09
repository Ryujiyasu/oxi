// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! Position edits keep formatting and follow Word's rules for joining and
//! splitting paragraphs.

use oxidocs_model::content::Hyperlink;
use oxidocs_model::edit::{self, apply, apply_all, Edit, Patch, Pos, Story};
use oxidocs_model::xml::Element;
use oxidocs_model::{Block, Cell, CellItem, DocxModel, Inline, Paragraph, Props, Row, RowItem, Run, RunItem, Table};

fn run(text: &str, props: &[Element]) -> Inline {
    let mut p = Props::new("w:rPr");
    for e in props {
        p.set(e.clone()).unwrap();
    }
    Inline::Run(Run { props: (!p.is_empty()).then_some(p), content: vec![RunItem::Text(text.into())] })
}

fn para(content: Vec<Inline>, jc: Option<&str>) -> Block {
    let props = jc.map(|v| {
        let mut p = Props::new("w:pPr");
        p.set(Element::new("w:jc").with_attr("w:val", v)).unwrap();
        p
    });
    Block::Paragraph(Paragraph { props, content, attrs: vec![] })
}

fn bold() -> Element {
    Element::new("w:b")
}

fn doc(body: Vec<Block>) -> DocxModel {
    DocxModel { body, ..Default::default() }
}

fn texts(m: &DocxModel) -> Vec<String> {
    edit::paragraphs(m, &Story::Body).unwrap().into_iter().map(edit::paragraph_text).collect()
}

/// Each paragraph as runs: (text, bold).
fn runs(m: &DocxModel, n: usize) -> Vec<(String, bool)> {
    let p = edit::paragraphs(m, &Story::Body).unwrap()[n];
    let mut out = Vec::new();
    fn walk(v: &[Inline], out: &mut Vec<(String, bool)>) {
        for i in v {
            match i {
                Inline::Run(r) => {
                    let b = r.props.as_ref().is_some_and(|p| p.get("w:b").is_some());
                    let t: String = r.content.iter().filter_map(|c| if let RunItem::Text(t) = c { Some(t.as_str()) } else { None }).collect();
                    out.push((t, b));
                }
                Inline::Hyperlink(h) => walk(&h.content, out),
                _ => {}
            }
        }
    }
    walk(&p.content, &mut out);
    out
}

fn jc(m: &DocxModel, n: usize) -> Option<String> {
    let p = edit::paragraphs(m, &Story::Body).unwrap()[n];
    p.props.as_ref()?.get("w:jc")?.attr("w:val").map(|s| s.to_string())
}

fn two_paras() -> DocxModel {
    doc(vec![
        para(vec![run("AA", &[bold()]), run("AA", &[])], Some("center")),
        para(vec![run("BBBB", &[])], Some("right")),
    ])
}

#[test]
fn typed_text_takes_the_formatting_before_it() {
    let mut m = two_paras();
    apply(&mut m, &Edit::InsertText { at: Pos::body(0, 2), text: "x".into() }).unwrap();
    assert_eq!(runs(&m, 0), vec![("AAx".into(), true), ("AA".into(), false)]);
    // At the start of a paragraph, the formatting after it.
    apply(&mut m, &Edit::InsertText { at: Pos::body(0, 0), text: "y".into() }).unwrap();
    assert_eq!(runs(&m, 0)[0], ("yAAx".into(), true));
    // Inside a run, the run grows.
    apply(&mut m, &Edit::InsertText { at: Pos::body(1, 2), text: "z\tz".into() }).unwrap();
    assert_eq!(texts(&m)[1], "BBz\tzBB");
}

#[test]
fn typed_text_beside_a_hyperlink_stays_outside_it() {
    let mut styled = Props::new("w:rPr");
    styled.set(Element::new("w:rStyle").with_attr("w:val", "Hyperlink")).unwrap();
    let link = Inline::Hyperlink(Hyperlink {
        target: Some("http://example.com/".into()),
        attrs: vec![("w:history".into(), "1".into())],
        content: vec![Inline::Run(Run { props: Some(styled), content: vec![RunItem::Text("AB".into())] })],
    });
    let mut m = doc(vec![para(vec![run("link:", &[]), link], None)]);
    let t = &texts(&m)[0];
    // 5 + begin + " HYPERLINK \"http://example.com/\"" + separator + AB + end, as Word counts.
    assert_eq!(t.chars().count(), 42);
    apply(&mut m, &Edit::InsertText { at: Pos::body(0, 42), text: "x".into() }).unwrap();
    let p = &m.body[0];
    let Block::Paragraph(p) = p else { panic!() };
    assert!(matches!(p.content.last(), Some(Inline::Run(r)) if r.props.is_none()), "typed run is outside the link and unstyled");
    // Inside the link's text, the link grows.
    apply(&mut m, &Edit::InsertText { at: Pos::body(0, 40), text: "y".into() }).unwrap();
    assert_eq!(runs(&m, 0)[1].0, "AyB");
    // Inside the field code: refused.
    assert!(apply(&mut m, &Edit::InsertText { at: Pos::body(0, 8), text: "q".into() }).is_err());
}

#[test]
fn positions_count_utf16_units_as_word_does() {
    // 📞 is outside the BMP: two positions in Word (and in JavaScript).
    let mut m = doc(vec![para(vec![run("a📞b", &[])], None)]);
    apply(&mut m, &Edit::InsertText { at: Pos::body(0, 3), text: "x".into() }).unwrap();
    assert_eq!(texts(&m)[0], "a📞xb");
    // Between the halves of the pair: refused.
    assert!(apply(&mut m, &Edit::InsertText { at: Pos::body(0, 2), text: "y".into() }).is_err());
    apply(&mut m, &Edit::Delete { from: Pos::body(0, 1), to: Pos::body(0, 3) }).unwrap();
    assert_eq!(texts(&m)[0], "axb");
}

#[test]
fn shown_characters_skip_field_codes() {
    let text = "a\u{13} HYPERLINK \"x\" \u{14}bc\u{15}\u{13} PAGE \u{14}1\u{15}\u{1c}d\u{1d}";
    let shown: String = edit::shown_positions(text).iter().map(|(c, _)| *c).collect();
    assert_eq!(shown, "abc1d");
    let at: Vec<usize> = edit::shown_positions(text).iter().map(|(_, p)| *p).collect();
    // a=0, begin=1, the 15-character code, separator=17, b=18, c=19, end=20,
    // begin=21, " PAGE "=22..27, separator=28, 1=29, end=30, start=31, d=32.
    assert_eq!(at, vec![0, 18, 19, 29, 32]);
}

#[test]
fn delete_inside_a_paragraph_keeps_the_runs_around_it() {
    let mut m = two_paras();
    apply(&mut m, &Edit::Delete { from: Pos::body(0, 1), to: Pos::body(0, 3) }).unwrap();
    assert_eq!(runs(&m, 0), vec![("A".into(), true), ("A".into(), false)]);
}

#[test]
fn delete_across_paragraphs_keeps_head_and_tail_apart() {
    // Word: "AA[AA¶BB]BB" deleted → "AA¶" (centred) and "BB¶" (right).
    let mut m = two_paras();
    apply(&mut m, &Edit::Delete { from: Pos::body(0, 2), to: Pos::body(1, 2) }).unwrap();
    assert_eq!(texts(&m), vec!["AA", "BB"]);
    assert_eq!((jc(&m, 0).as_deref(), jc(&m, 1).as_deref()), (Some("center"), Some("right")));
}

#[test]
fn delete_from_a_paragraph_start_removes_it() {
    let mut m = two_paras();
    apply(&mut m, &Edit::Delete { from: Pos::body(0, 0), to: Pos::body(1, 2) }).unwrap();
    assert_eq!(texts(&m), vec!["BB"]);
    assert_eq!(jc(&m, 0).as_deref(), Some("right"));
}

#[test]
fn delete_ending_at_a_paragraph_start_joins() {
    // The range takes in the first paragraph's mark only.
    let mut m = two_paras();
    apply(&mut m, &Edit::Delete { from: Pos::body(0, 4), to: Pos::body(1, 0) }).unwrap();
    assert_eq!(texts(&m), vec!["AAAABBBB"]);
    assert_eq!(jc(&m, 0).as_deref(), Some("center"));
}

#[test]
fn typing_over_paragraphs_joins_with_the_first_ones_formatting() {
    let mut m = two_paras();
    apply(&mut m, &Edit::Replace { from: Pos::body(0, 2), to: Pos::body(1, 2), text: "X".into() }).unwrap();
    assert_eq!(texts(&m), vec!["AAXBB"]);
    assert_eq!(jc(&m, 0).as_deref(), Some("center"));
    // The typed text takes the formatting of the range's first character.
    assert_eq!(runs(&m, 0), vec![("AA".into(), true), ("XBB".into(), false)]);
}

#[test]
fn typing_over_a_range_takes_its_first_characters_formatting() {
    let mut m = two_paras();
    apply(&mut m, &Edit::Replace { from: Pos::body(0, 1), to: Pos::body(0, 3), text: "xy".into() }).unwrap();
    assert_eq!(runs(&m, 0), vec![("Axy".into(), true), ("A".into(), false)]);
}

#[test]
fn enter_gives_both_halves_the_paragraph_formatting() {
    let mut m = two_paras();
    apply(&mut m, &Edit::SplitParagraph { at: Pos::body(0, 3) }).unwrap();
    assert_eq!(texts(&m), vec!["AAA", "A", "BBBB"]);
    assert_eq!(runs(&m, 0), vec![("AA".into(), true), ("A".into(), false)]);
    assert_eq!(runs(&m, 1), vec![("A".into(), false)]);
    assert_eq!((jc(&m, 0).as_deref(), jc(&m, 1).as_deref()), (Some("center"), Some("center")));
}

#[test]
fn join_keeps_the_first_paragraphs_formatting() {
    let mut m = two_paras();
    apply(&mut m, &Edit::JoinParagraphs { story: Story::Body, para: 0 }).unwrap();
    assert_eq!(texts(&m), vec!["AAAABBBB"]);
    assert_eq!(jc(&m, 0).as_deref(), Some("center"));
}

#[test]
fn formatting_a_range_splits_runs_at_its_edges() {
    let mut m = two_paras();
    let patch = Patch { set: vec![bold()], ..Default::default() };
    apply(&mut m, &Edit::FormatText { from: Pos::body(1, 1), to: Pos::body(1, 3), patch }).unwrap();
    assert_eq!(runs(&m, 1), vec![("B".into(), false), ("BB".into(), true), ("B".into(), false)]);
    let unbold = Patch { remove: vec!["w:b".into()], ..Default::default() };
    apply(&mut m, &Edit::FormatText { from: Pos::body(0, 0), to: Pos::body(0, 1), patch: unbold }).unwrap();
    assert_eq!(runs(&m, 0), vec![("A".into(), false), ("A".into(), true), ("AA".into(), false)]);
}

#[test]
fn formatting_past_a_paragraph_end_formats_its_mark() {
    let mut m = two_paras();
    let patch = Patch { set: vec![bold()], ..Default::default() };
    apply(&mut m, &Edit::FormatText { from: Pos::body(0, 2), to: Pos::body(1, 1), patch }).unwrap();
    let Block::Paragraph(p) = &m.body[0] else { panic!() };
    assert!(p.props.as_ref().unwrap().get("w:rPr").unwrap().child("w:b").is_some());
}

#[test]
fn paragraph_formatting_is_set_in_schema_order() {
    let mut m = two_paras();
    let patch = Patch {
        set: vec![Element::new("w:spacing").with_attr("w:after", "200"), Element::new("w:jc").with_attr("w:val", "both")],
        ..Default::default()
    };
    apply(&mut m, &Edit::FormatParagraphs { from: Pos::body(0, 0), to: Pos::body(1, 0), patch }).unwrap();
    assert_eq!(jc(&m, 1).as_deref(), Some("both"));
    // Merging sets attributes and keeps the others.
    let merge = Patch { merge: vec![Element::new("w:spacing").with_attr("w:line", "360").with_attr("w:lineRule", "auto")], ..Default::default() };
    apply(&mut m, &Edit::FormatParagraphs { from: Pos::body(0, 0), to: Pos::body(0, 0), patch: merge }).unwrap();
    let p = edit::paragraphs(&m, &Story::Body).unwrap()[0];
    let spacing = p.props.as_ref().unwrap().get("w:spacing").unwrap();
    assert_eq!((spacing.attr("w:after"), spacing.attr("w:line")), (Some("200"), Some("360")));
    // A property the schema does not allow is refused, not written.
    let bad = Patch { set: vec![Element::new("w:nonsense")], ..Default::default() };
    assert!(apply(&mut m, &Edit::FormatParagraphs { from: Pos::body(0, 0), to: Pos::body(0, 0), patch: bad }).is_err());
}

#[test]
fn table_paragraphs_are_numbered_and_a_trailing_table_gets_a_final_paragraph() {
    let cell = |t: &str| CellItem::Cell(Cell { props: None, content: vec![para(vec![run(t, &[])], None)], attrs: vec![] });
    let row = RowItem::Row(Row { props_ex: None, props: None, cells: vec![cell("c1"), cell("c2")], attrs: vec![], props_after_cells: false });
    let table = Block::Table(Table { props: Props::new("w:tblPr"), grid: None, rows: vec![row] });
    let mut m = doc(vec![para(vec![run("before", &[])], None), table]);
    assert_eq!(texts(&m), vec!["before", "c1", "c2", ""]);
    apply(&mut m, &Edit::InsertText { at: Pos::body(2, 2), text: "!".into() }).unwrap();
    apply(&mut m, &Edit::InsertText { at: Pos::body(3, 0), text: "after".into() }).unwrap();
    assert_eq!(texts(&m), vec!["before", "c1", "c2!", "after"]);
    // Paragraphs in different containers do not join.
    assert!(apply(&mut m, &Edit::Delete { from: Pos::body(0, 2), to: Pos::body(1, 1) }).is_err());
}

#[test]
fn a_cell_without_a_paragraph_has_one() {
    let bare = CellItem::Cell(Cell { props: None, content: vec![], attrs: vec![] });
    let full = CellItem::Cell(Cell { props: None, content: vec![para(vec![run("c2", &[])], None)], attrs: vec![] });
    let row = RowItem::Row(Row { props_ex: None, props: None, cells: vec![bare, full], attrs: vec![], props_after_cells: false });
    let table = Block::Table(Table { props: Props::new("w:tblPr"), grid: None, rows: vec![row] });
    let mut m = doc(vec![table, para(vec![run("after", &[])], None)]);
    assert_eq!(texts(&m), vec!["", "c2", "after"]);
    apply(&mut m, &Edit::InsertText { at: Pos::body(0, 0), text: "c1".into() }).unwrap();
    assert_eq!(texts(&m), vec!["c1", "c2", "after"]);
}

#[test]
fn a_content_control_s_start_and_end_take_positions() {
    use oxidocs_model::content::Sdt;
    let control = Block::Sdt(Sdt { props: None, end_props: None, content: vec![para(vec![run("in", &[])], None)] });
    let mut m = doc(vec![control, para(vec![run("out", &[])], None)]);
    assert_eq!(edit::paragraph_texts(&m, &Story::Body).unwrap(), vec!["\u{1c}in\u{1d}", "out"]);
    // Offset 1 is the start of the control's text.
    apply(&mut m, &Edit::InsertText { at: Pos::body(0, 1), text: ">".into() }).unwrap();
    assert_eq!(texts(&m)[0], ">in");
}

#[test]
fn a_failed_batch_leaves_the_model_alone() {
    let mut m = two_paras();
    let before = m.clone();
    let edits = [
        Edit::InsertText { at: Pos::body(0, 0), text: "ok".into() },
        Edit::InsertText { at: Pos::body(0, 99), text: "past the end".into() },
    ];
    assert!(apply_all(&mut m, &edits).is_err());
    assert_eq!(m, before);
}

#[test]
fn edits_read_from_json() {
    let json = r#"[
        {"op": "insert_text", "at": {"para": 0, "offset": 0}, "text": "Hi "},
        {"op": "format_text", "from": {"para": 1, "offset": 0}, "to": {"para": 1, "offset": 2},
         "patch": {"set": [{"name": "w:i"}]}},
        {"op": "split_paragraph", "at": {"story": "body", "para": 1, "offset": 2}}
    ]"#;
    let edits: Vec<Edit> = serde_json::from_str(json).unwrap();
    let mut m = two_paras();
    apply_all(&mut m, &edits).unwrap();
    assert_eq!(texts(&m), vec!["Hi AAAA", "BB", "BB"]);
    // The result is a document Word's schema order accepts.
    let bytes = m.write(None).unwrap();
    let back = DocxModel::read(&bytes).unwrap();
    assert_eq!(texts(&back), vec!["Hi AAAA", "BB", "BB"]);
}
