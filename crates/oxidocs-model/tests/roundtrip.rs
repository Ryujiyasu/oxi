// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! A document built from data survives writing and reading back, and the
//! canonical comparison notices a change.

use oxidocs_model::content::{Hyperlink, Object, Part};
use oxidocs_model::package::{HeaderFooter, Note};
use oxidocs_model::xml::{parse, Element};
use oxidocs_model::{canon, Block, Cell, CellItem, DocxModel, Inline, Package, Paragraph, Props, Row, RowItem, Run, RunItem, Table};

fn run(text: &str, props: &[Element]) -> Inline {
    let mut p = Props::new("w:rPr");
    for e in props {
        p.set(e.clone()).unwrap();
    }
    Inline::Run(Run { props: (!p.is_empty()).then_some(p), content: vec![RunItem::Text(text.into())] })
}

fn para(content: Vec<Inline>, props: &[Element]) -> Block {
    let mut p = Props::new("w:pPr");
    for e in props {
        p.set(e.clone()).unwrap();
    }
    Block::Paragraph(Paragraph { props: (!p.is_empty()).then_some(p), content, attrs: vec![] })
}

const PNG: &[u8] = &[0x89, b'P', b'N', b'G', 13, 10, 26, 10, 0, 0, 0, 0];

fn sample() -> DocxModel {
    let picture = parse(r##"<w:drawing xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><wp:inline><wp:extent cx="914400" cy="914400"/><wp:docPr id="1" name="p"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:blipFill><a:blip r:embed="#0"/></pic:blipFill></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing>"##).unwrap();
    let image = Part::Internal {
        rel_type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image".into(),
        name: "media/image1.png".into(),
        content_type: "image/png".into(),
        data: PNG.to_vec(),
        rels: vec![],
    };
    let cell = |t: &str| CellItem::Cell(Cell { props: None, content: vec![para(vec![run(t, &[])], &[])], attrs: vec![] });
    let mut tbl = Props::new("w:tblPr");
    tbl.set(Element::new("w:tblW").with_attr("w:w", "0").with_attr("w:type", "auto")).unwrap();
    let mut trpr = Props::new("w:trPr");
    trpr.set(Element::new("w:trHeight").with_attr("w:val", "400").with_attr("w:hRule", "exact")).unwrap();
    let mut section = Props::new("w:sectPr");
    section.push(Element::new("w:headerReference").with_attr("w:type", "default").with_attr("r:id", "#0")).unwrap();
    section.set(Element::new("w:pgSz").with_attr("w:w", "11906").with_attr("w:h", "16838")).unwrap();
    DocxModel {
        background: None,
        body: vec![
            para(
                vec![
                    run("見出し", &[Element::new("w:b"), Element::new("w:sz").with_attr("w:val", "32")]),
                    Inline::Hyperlink(Hyperlink {
                        target: Some("https://example.com/".into()),
                        attrs: vec![],
                        content: vec![run(" link", &[])],
                    }),
                    Inline::Run(Run {
                        props: None,
                        content: vec![RunItem::Mark { name: "w:footnoteReference".into(), attrs: vec![("w:id".into(), "1".into())] }],
                    }),
                ],
                &[Element::new("w:jc").with_attr("w:val", "center"), Element::new("w:ind").with_attr("w:firstLineChars", "100")],
            ),
            Block::Table(Table {
                props: tbl,
                grid: Element::new("w:tblGrid").with_child(Element::new("w:gridCol").with_attr("w:w", "4000")),
                rows: vec![RowItem::Row(Row { props_ex: None, props: Some(trpr), cells: vec![cell("A")], attrs: vec![] })],
            }),
            Block::Paragraph(Paragraph {
                props: None,
                content: vec![Inline::Run(Run { props: None, content: vec![RunItem::Object(Object { xml: picture, parts: vec![image] })] })],
                attrs: vec![],
            }),
        ],
        section: Some(section),
        headers_footers: vec![HeaderFooter { kind: "w:hdr".into(), content: vec![para(vec![run("header", &[])], &[])] }],
        footnotes: Some(vec![
            Note { attrs: vec![("w:type".into(), "separator".into()), ("w:id".into(), "-1".into())], content: vec![para(vec![], &[])] },
            Note { attrs: vec![("w:id".into(), "1".into())], content: vec![para(vec![run("note", &[])], &[])] },
        ]),
        endnotes: None,
    }
}

#[test]
fn a_built_document_reads_back_as_the_same_model() {
    let model = sample();
    let bytes = model.write(None).unwrap();
    let back = DocxModel::read(&bytes).unwrap();
    assert_eq!(back, model);
}

#[test]
fn the_comparison_sees_a_one_letter_change() {
    let model = sample();
    let a = Package::read(&model.write(None).unwrap()).unwrap();
    assert!(canon::compare(&a, &a).is_empty());
    let mut changed = model.clone();
    if let Some(Block::Paragraph(p)) = changed.headers_footers[0].content.first_mut() {
        if let Some(Inline::Run(r)) = p.content.first_mut() {
            r.content = vec![RunItem::Text("Header".into())];
        }
    }
    let b = Package::read(&changed.write(None).unwrap()).unwrap();
    let diffs = canon::compare(&a, &b);
    assert!(!diffs.is_empty());
}

#[test]
fn json_is_the_creation_format() {
    let model = sample();
    let json = serde_json::to_string(&model).unwrap();
    let from_json: DocxModel = serde_json::from_str(&json).unwrap();
    assert_eq!(from_json, model);
}
