// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! What survives a round trip through the editor, for the styling a cell wears.

use oxicells_core::editor::XlsxEditor;
use oxicells_core::ir::{BorderLine, CellStyle};
use oxicells_core::parser::parse_xlsx;

const FIXTURE: &[u8] = include_bytes!("fixtures/hidden_rows_cols.xlsx");

fn style_at(
    workbook: &oxicells_core::ir::Workbook,
    row: u32,
    col: u32,
) -> CellStyle {
    workbook.sheets[0]
        .rows
        .iter()
        .find(|held| held.index == row)
        .and_then(|held| held.cells.iter().find(|cell| cell.col == col))
        .map(|cell| cell.style.clone())
        .unwrap_or_default()
}

#[test]
fn a_style_the_editor_sets_is_saved() {
    let mut editor = XlsxEditor::new(FIXTURE).expect("the fixture opens");
    editor.set_cell_style(
        0,
        1,
        0,
        CellStyle {
            bold: true,
            italic: true,
            font_size: Some(14.0),
            ..CellStyle::default()
        },
    );
    let saved = editor.save().expect("the workbook saves");

    let workbook = parse_xlsx(&saved).expect("the saved workbook parses");
    let style = style_at(&workbook, 1, 0);
    assert!(style.bold);
    assert!(style.italic);
    assert_eq!(style.font_size, Some(14.0));
}

#[test]
fn colours_borders_and_alignment_are_saved() {
    let mut editor = XlsxEditor::new(FIXTURE).expect("the fixture opens");
    editor.set_cell_style(
        0,
        1,
        0,
        CellStyle {
            font_color: Some("FF0000".to_string()),
            bg_color: Some("FFFF00".to_string()),
            horizontal_align: Some("center".to_string()),
            border_top: Some(BorderLine { style: "thin".to_string(), color: None }),
            border_bottom: Some(BorderLine { style: "medium".to_string(), color: None }),
            ..CellStyle::default()
        },
    );
    let saved = editor.save().expect("the workbook saves");

    let workbook = parse_xlsx(&saved).expect("the saved workbook parses");
    let style = style_at(&workbook, 1, 0);
    assert_eq!(style.font_color.as_deref(), Some("FF0000"));
    assert_eq!(style.bg_color.as_deref(), Some("FFFF00"));
    assert_eq!(style.horizontal_align.as_deref(), Some("center"));
    assert!(style.border_top.is_some());
    assert!(style.border_bottom.is_some());
    assert!(style.border_left.is_none());
}

#[test]
fn a_number_format_is_saved() {
    let mut editor = XlsxEditor::new(FIXTURE).expect("the fixture opens");
    editor.set_cell_style(
        0,
        1,
        0,
        CellStyle {
            number_format: Some("0.000".to_string()),
            ..CellStyle::default()
        },
    );
    let saved = editor.save().expect("the workbook saves");

    let workbook = parse_xlsx(&saved).expect("the saved workbook parses");
    assert_eq!(
        style_at(&workbook, 1, 0).number_format.as_deref(),
        Some("0.000")
    );
}

/// Cells that share a style share one entry in the style sheet.
#[test]
fn cells_wearing_the_same_style_share_one_entry() {
    let bold = CellStyle {
        bold: true,
        ..CellStyle::default()
    };
    let mut editor = XlsxEditor::new(FIXTURE).expect("the fixture opens");
    editor.set_cell_style(0, 1, 0, bold.clone());
    editor.set_cell_style(0, 3, 0, bold.clone());
    editor.set_cell_style(0, 5, 0, bold);
    let saved = editor.save().expect("the workbook saves");

    let workbook = parse_xlsx(&saved).expect("the saved workbook parses");
    for row in [1, 3, 5] {
        assert!(style_at(&workbook, row, 0).bold, "row {row} is bold");
    }

    // One xf, not three: the style sheet grew by a single entry.
    let mut archive = zip::ZipArchive::new(std::io::Cursor::new(&saved)).expect("opens");
    let mut xml = String::new();
    std::io::Read::read_to_string(
        &mut archive.by_name("xl/styles.xml").expect("has styles"),
        &mut xml,
    )
    .expect("reads");
    assert_eq!(xml.matches("<b/>").count(), 1);
}

/// The whole way round, as a VBA run would do it.
#[test]
fn a_changed_workbook_saves_the_styling_it_changed() {
    let mut workbook = parse_xlsx(FIXTURE).expect("the fixture parses");
    {
        let first = workbook.sheets[0]
            .rows
            .iter_mut()
            .find(|row| row.index == 1)
            .expect("row 1");
        first.cells[0].style.bold = true;
    }

    let mut editor = XlsxEditor::new(FIXTURE).expect("the fixture opens");
    editor
        .apply_workbook(&workbook)
        .expect("the change is one the editor can write");
    let saved = editor.save().expect("the workbook saves");

    let reread = parse_xlsx(&saved).expect("the saved workbook parses");
    assert!(style_at(&reread, 1, 0).bold);
    // The rows it hides are still hidden.
    let hidden: Vec<u32> = reread.sheets[0]
        .rows
        .iter()
        .filter(|row| row.hidden)
        .map(|row| row.index)
        .collect();
    assert_eq!(hidden, vec![2, 4]);
}

#[test]
fn an_unstyled_workbook_writes_nothing_back() {
    let workbook = parse_xlsx(FIXTURE).expect("the fixture parses");
    let mut editor = XlsxEditor::new(FIXTURE).expect("the fixture opens");
    editor.apply_workbook(&workbook).expect("nothing to write");
    assert!(!editor.has_edits());
}


#[test]
fn the_typeface_a_cell_names_is_read_and_saved() {
    // The fixture was authored in 游ゴシック, so every cell already names a
    // typeface before the editor touches anything.
    let opened = parse_xlsx(FIXTURE).expect("the fixture parses");
    assert_eq!(
        style_at(&opened, 1, 0).font_name.as_deref(),
        Some("游ゴシック")
    );

    let mut editor = XlsxEditor::new(FIXTURE).expect("the fixture opens");
    editor.set_cell_style(
        0,
        1,
        0,
        CellStyle {
            font_name: Some("Meiryo UI".to_string()),
            font_size: Some(9.0),
            ..CellStyle::default()
        },
    );
    let saved = editor.save().expect("the workbook saves");

    let workbook = parse_xlsx(&saved).expect("the saved workbook parses");
    let style = style_at(&workbook, 1, 0);
    assert_eq!(style.font_name.as_deref(), Some("Meiryo UI"));
    assert_eq!(style.font_size, Some(9.0));
}

#[test]
fn how_a_cell_places_and_breaks_its_text_is_read() {
    // A statistics table from the corpus: every cell centres its text
    // vertically, and the header cells break onto a second line.
    const TABLE: &[u8] =
        include_bytes!("../../../tools/golden-test/documents/xlsx/16c7b9f9ed53_toukeihyo.xlsx");
    let workbook = parse_xlsx(TABLE).expect("the workbook parses");

    let placed = workbook.sheets[0]
        .rows
        .iter()
        .flat_map(|row| &row.cells)
        .filter(|cell| cell.style.vertical_align.is_some())
        .count();
    assert!(placed > 0, "no cell said where its text sits");

    let wrapped = workbook.sheets[0]
        .rows
        .iter()
        .flat_map(|row| &row.cells)
        .filter(|cell| cell.style.wrap_text)
        .count();
    assert!(wrapped > 0, "no cell said it breaks its text");
}

#[test]
fn a_cell_wears_the_font_of_the_style_it_is_built_on() {
    // The hyperlinks in this workbook name no font of their own: their format
    // points at the built-in Hyperlink style, whose font is underlined. Reading
    // only the cell's own format loses that.
    const LINKED: &[u8] =
        include_bytes!("../../../tools/golden-test/documents/xlsx/24d76e2a8663_h2daa202505_jikei.xlsx");
    let workbook = parse_xlsx(LINKED).expect("the workbook parses");

    let underlined = workbook.sheets[0]
        .rows
        .iter()
        .flat_map(|row| &row.cells)
        .filter(|cell| cell.style.underline)
        .count();
    assert!(underlined > 0, "no cell inherited the underlined link font");
}

#[test]
fn a_colour_the_theme_names_is_resolved() {
    // The hyperlinks in this workbook state no colour of their own: their font
    // names theme colour 10, which the Office theme calls hlink and states as
    // 0563C1. Excel paints them that blue.
    const LINKED: &[u8] =
        include_bytes!("../../../tools/golden-test/documents/xlsx/24d76e2a8663_h2daa202505_jikei.xlsx");
    let workbook = parse_xlsx(LINKED).expect("the workbook parses");

    let blue = workbook.sheets[0]
        .rows
        .iter()
        .flat_map(|row| &row.cells)
        .filter(|cell| {
            cell.style.underline && cell.style.font_color.as_deref() == Some("0563C1")
        })
        .count();
    assert!(blue > 0, "no link came out the theme's link blue");
}

#[test]
fn a_sheet_says_how_far_it_reaches() {
    // This sheet declares B2:F149 while its last filled cell stops short of
    // column F. Excel hands the declared range over when asked for a picture,
    // and text running on past the last cell needs that room to land in.
    const FORM: &[u8] =
        include_bytes!("../../../tools/golden-test/documents/xlsx/14eaa2188aca_001904853.xlsx");
    let workbook = parse_xlsx(FORM).expect("the workbook parses");
    assert_eq!(workbook.sheets[0].declared_range, Some((2, 1, 149, 5)));
}

#[test]
fn a_table_is_read_with_the_colour_its_style_dresses_it_in() {
    // TableStyleMedium7 is the seventh of the Medium family, which walk the
    // theme's accents from accent1; the seventh lands on accent6, which this
    // workbook's theme states as 4EA72E. Its banded rows are that under a tint.
    const TABLED: &[u8] =
        include_bytes!("../../../tools/golden-test/documents/xlsx/2b051dcf28c4_doi-list.xlsx");
    let workbook = parse_xlsx(TABLED).expect("the workbook parses");

    let table = workbook.sheets[0]
        .tables
        .first()
        .expect("the sheet holds a table");
    assert_eq!(table.style.as_deref(), Some("TableStyleMedium7"));
    assert_eq!(table.accent.as_deref(), Some("4EA72E"));
    assert_eq!(table.header_rows, 1);
    assert!(table.banded_rows);
    assert_eq!((table.start_row, table.start_col), (2, 0));
}
