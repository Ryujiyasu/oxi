// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! Opening a file works out the formulas it arrives without answers for, and
//! only those.
//!
//! `parse_xlsx` ends by filling those gaps. It used to do so by working out
//! EVERY formula and then throwing away all but the missing ones, which on a
//! 936,000-cell workbook was twenty of the twenty-two seconds it took to open —
//! and bought nothing at all, because a file Excel wrote has an answer for
//! every formula in it.
//!
//! What must not change is the filling itself, which is why it is pinned here
//! from both sides: a gap is filled, and an answer already in the file is left
//! exactly as it stands even when it is wrong. That second half is what makes
//! opening a file show what Excel showed rather than what this evaluator
//! thinks, and it is the reason the pass is `OnlyMissing` at all.

use oxicells_core::ir::CellValue;
use oxicells_core::{editor::CellEditValue, parse_xlsx, XlsxEditor};

const FIXTURE: &[u8] = include_bytes!("fixtures/frozen.xlsx");

/// A workbook whose first sheet holds two numbers and two formulas over them.
fn a_workbook() -> Vec<u8> {
    let mut editor = XlsxEditor::new(FIXTURE).expect("opens");
    editor.set_cell_value(0, 1, 0, CellEditValue::Number(10.0));
    editor.set_cell_value(0, 2, 0, CellEditValue::Number(32.0));
    editor.set_cell_value(0, 1, 1, CellEditValue::Formula("=A1+A2".to_string()));
    editor.set_cell_value(0, 2, 1, CellEditValue::Formula("=B1*2".to_string()));
    editor.save().expect("saves")
}

/// The same file with the cached answers rewritten: `gap` loses its `<v>`
/// entirely, and whatever `wrong` is given it keeps.
fn with_answers(source: &[u8], gap: &str, wrong: Option<(&str, &str)>) -> Vec<u8> {
    let mut out = Vec::new();
    {
        let reader = std::io::Cursor::new(source);
        let mut archive = zip::ZipArchive::new(reader).expect("a zip");
        let mut writer = zip::ZipWriter::new(std::io::Cursor::new(&mut out));
        for at in 0..archive.len() {
            let mut part = archive.by_index(at).expect("a part");
            let name = part.name().to_string();
            let mut bytes = Vec::new();
            std::io::copy(&mut part, &mut bytes).expect("reads");
            if name == "xl/worksheets/sheet1.xml" {
                let mut xml = String::from_utf8_lossy(&bytes).to_string();
                xml = drop_value(&xml, gap);
                if let Some((at, held)) = wrong {
                    xml = set_value(&xml, at, held);
                }
                bytes = xml.into_bytes();
            }
            writer
                .start_file(name, zip::write::SimpleFileOptions::default())
                .expect("starts");
            use std::io::Write;
            writer.write_all(&bytes).expect("writes");
        }
        writer.finish().expect("finishes");
    }
    out
}

/// Everything between the opening tag of cell `at` and its close.
fn body_of<'a>(xml: &'a str, at: &str) -> (usize, usize, &'a str) {
    let opening = format!("<c r=\"{at}\"");
    let start = xml.find(&opening).unwrap_or_else(|| panic!("no cell {at}"));
    let end = xml[start..].find("</c>").expect("a closing tag") + start;
    (start, end, &xml[start..end])
}

fn drop_value(xml: &str, at: &str) -> String {
    if !xml.contains(&format!("<c r=\"{at}\"")) {
        return xml.to_string();
    }
    let (start, end, body) = body_of(xml, at);
    let without = match (body.find("<v>"), body.find("</v>")) {
        (Some(from), Some(to)) => format!("{}{}", &body[..from], &body[to + 4..]),
        _ => body.to_string(),
    };
    format!("{}{}{}", &xml[..start], without, &xml[end..])
}

fn set_value(xml: &str, at: &str, held: &str) -> String {
    let (start, end, body) = body_of(xml, at);
    let (from, to) = (
        body.find("<v>").expect("a value"),
        body.find("</v>").expect("a value"),
    );
    let replaced = format!("{}<v>{held}</v>{}", &body[..from], &body[to + 4..]);
    format!("{}{}{}", &xml[..start], replaced, &xml[end..])
}

fn shown(book: &oxicells_core::ir::Workbook, row: u32, col: u32) -> CellValue {
    book.sheets[0]
        .rows
        .iter()
        .find(|one| one.index == row)
        .and_then(|one| one.cells.iter().find(|cell| cell.col == col))
        .map(|cell| cell.value.clone())
        .unwrap_or(CellValue::Empty)
}

#[test]
fn a_formula_that_arrives_without_an_answer_is_worked_out() {
    // B1 loses its cached value; B2 keeps one that reads B1.
    let book = parse_xlsx(&with_answers(&a_workbook(), "B1", None)).expect("parses");
    assert_eq!(shown(&book, 1, 1).display(), "42", "10 + 32");
}

#[test]
fn an_answer_the_file_already_holds_is_left_alone_even_when_it_is_wrong() {
    // B2 says 999 where `=B1*2` would say 84. Opening the file must show what
    // the file says: this pass fills gaps, it does not audit.
    //
    // It is also what makes the skip safe — a workbook with no gaps needs no
    // pass at all, and this is the assertion that would fail if "no gaps"
    // were ever taken to mean "recalculate everything".
    let book = parse_xlsx(&with_answers(&a_workbook(), "B1", Some(("B2", "999"))))
        .expect("parses");
    assert_eq!(shown(&book, 1, 1).display(), "42", "the gap was filled");
    assert_eq!(shown(&book, 2, 1).display(), "999", "and the answer kept");
}

#[test]
fn a_workbook_with_no_gaps_comes_through_untouched() {
    // Every formula answered, one of them wrongly. Nothing is worked out, so
    // the wrong one stays wrong — which is the whole point of opening a file
    // showing what Excel showed.
    let book = parse_xlsx(&with_answers(&a_workbook(), "Z9", Some(("B2", "999"))))
        .expect("parses");
    assert_eq!(shown(&book, 2, 1).display(), "999");
    // And the formulas are all still there to be worked out on demand.
    let mut book = book;
    oxicells_core::formula::evaluate_workbook_formulas(&mut book);
    assert_eq!(shown(&book, 2, 1).display(), "84", "asked properly, 42 * 2");
}

#[test]
fn a_gap_whose_input_is_also_a_gap_is_worked_out_in_order() {
    // B2 reads B1, and neither has an answer. Filling only the missing cells
    // is correct because a missing input comes earlier in the same order —
    // this is the case that would break if they were filled in sheet order.
    let mut source = with_answers(&a_workbook(), "B1", None);
    source = with_answers(&source, "B2", None);
    let book = parse_xlsx(&source).expect("parses");
    assert_eq!(shown(&book, 1, 1).display(), "42");
    assert_eq!(shown(&book, 2, 1).display(), "84");
}
