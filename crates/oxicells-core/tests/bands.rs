// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! Putting a row in and taking one out.
//!
//! The operation touches more of a workbook than any other, and every part of
//! it fails quietly: a merge left behind, a frozen row off by one, a formula
//! still pointing where the number used to be. So each thing that has to move
//! is asked about separately here rather than trusted to a single "it works".

use oxicells_core::bands::{self, Band};
use oxicells_core::ir::{Cell, CellValue, MergeCell, Row, Sheet, Table, Workbook};

/// A sheet holding, in rows 1 to 4, the numbers 1..4 in column A and a running
/// formula in column B.
fn a_sheet(name: &str) -> Sheet {
    let mut sheet = Sheet {
        name: name.to_string(),
        rows: Vec::new(),
        col_count: 4,
        col_widths: vec![8.0, 9.0, 10.0, 11.0],
        default_col_width: 8.43,
        default_row_height: 15.0,
        default_row_custom: false,
        col_fonts: Vec::new(),
        normal_font: None,
        first_font: None,
        frozen_rows: 0,
        frozen_cols: 0,
        merge_cells: Vec::new(),
        hidden_cols: Vec::new(),
        declared_range: None,
        tables: Vec::new(),
        drawings: Vec::new(),
        comments: Vec::new(),
        auto_filter: None,
        unsupported_elements: Vec::new(),
    };
    for index in 1..=4u32 {
        sheet.rows.push(Row {
            index,
            cells: vec![
                Cell {
                    col: 0,
                    value: CellValue::Number(index as f64),
                    style: Default::default(),
                    formula: None,
                    runs: Vec::new(),
                },
                Cell {
                    col: 1,
                    value: CellValue::Empty,
                    style: Default::default(),
                    formula: Some(format!("=A{index}*10")),
                    runs: Vec::new(),
                },
            ],
            height: Some(15.0 + index as f32),
            custom_height: true,
            style_font: None,
            thick_top: false,
            thick_bottom: false,
            hidden: false,
        });
    }
    sheet
}

fn a_workbook() -> Workbook {
    let mut book = Workbook::default();
    book.sheets.push(a_sheet("Sheet1"));
    book
}

fn formula_at(book: &Workbook, sheet: &str, row: u32, col: u32) -> Option<String> {
    book.sheets
        .iter()
        .find(|one| one.name == sheet)?
        .rows
        .iter()
        .find(|one| one.index == row)?
        .cells
        .iter()
        .find(|cell| cell.col == col)?
        .formula
        .clone()
}

fn number_at(book: &Workbook, sheet: &str, row: u32, col: u32) -> Option<f64> {
    match book
        .sheets
        .iter()
        .find(|one| one.name == sheet)?
        .rows
        .iter()
        .find(|one| one.index == row)?
        .cells
        .iter()
        .find(|cell| cell.col == col)?
        .value
    {
        CellValue::Number(held) => Some(held),
        _ => None,
    }
}

#[test]
fn an_inserted_row_pushes_the_rows_below_it_down() {
    let mut book = a_workbook();
    bands::insert(&mut book, "Sheet1", Band::Rows, 2, 1);
    assert_eq!(number_at(&book, "Sheet1", 1, 0), Some(1.0), "row 1 stayed");
    assert_eq!(number_at(&book, "Sheet1", 2, 0), None, "row 2 is now empty");
    assert_eq!(number_at(&book, "Sheet1", 3, 0), Some(2.0), "what was row 2");
    assert_eq!(number_at(&book, "Sheet1", 5, 0), Some(4.0), "what was row 4");
    // A row's height belongs to the row, so it travels with it.
    let moved = book.sheets[0].rows.iter().find(|one| one.index == 3).unwrap();
    assert_eq!(moved.height, Some(17.0), "row 2's own height, at row 3");
}

#[test]
fn the_formulas_follow_the_rows_they_name() {
    let mut book = a_workbook();
    bands::insert(&mut book, "Sheet1", Band::Rows, 2, 1);
    assert_eq!(formula_at(&book, "Sheet1", 1, 1).as_deref(), Some("=A1*10"));
    assert_eq!(
        formula_at(&book, "Sheet1", 3, 1).as_deref(),
        Some("=A3*10"),
        "the formula that was in row 2 reading A2 now sits in row 3 reading A3",
    );
    assert_eq!(formula_at(&book, "Sheet1", 5, 1).as_deref(), Some("=A5*10"));
}

#[test]
fn a_formula_reading_a_row_that_was_taken_out_says_so() {
    // Excel writes `#REF!` into the formula itself, because shifting the
    // reference along would answer confidently with somebody else's number.
    let mut book = a_workbook();
    bands::remove(&mut book, "Sheet1", Band::Rows, 2, 1);
    assert_eq!(number_at(&book, "Sheet1", 2, 0), Some(3.0), "row 3 came up");
    let orphaned = book.sheets[0]
        .rows
        .iter()
        .flat_map(|row| row.cells.iter())
        .filter_map(|cell| cell.formula.clone())
        .find(|held| held.contains("#REF!"));
    assert_eq!(orphaned, None, "the formula that read A2 went with its row");
    // And the ones that survived came up with their rows.
    assert_eq!(formula_at(&book, "Sheet1", 2, 1).as_deref(), Some("=A2*10"));
    assert_eq!(formula_at(&book, "Sheet1", 3, 1).as_deref(), Some("=A3*10"));
}

#[test]
fn a_formula_on_another_sheet_follows_a_row_it_names() {
    // `=Sheet1!A3` is not on Sheet1, but it is about Sheet1.
    let mut book = a_workbook();
    let mut other = a_sheet("Summary");
    other.rows[0].cells[1].formula = Some("=Sheet1!A3+Summary!A3".to_string());
    book.sheets.push(other);
    bands::insert(&mut book, "Sheet1", Band::Rows, 2, 1);
    assert_eq!(
        formula_at(&book, "Summary", 1, 1).as_deref(),
        Some("=Sheet1!A4+Summary!A3"),
        "the reference to Sheet1 moved and the one to Summary did not",
    );
}

#[test]
fn a_merge_the_row_lands_inside_grows_and_one_below_it_slides() {
    let mut book = a_workbook();
    book.sheets[0].merge_cells = vec![
        // Rows 1 to 3: an insert at row 2 lands inside it.
        MergeCell { start_row: 1, start_col: 0, end_row: 3, end_col: 0 },
        // Row 4 alone across two columns: below the insert, so it slides.
        MergeCell { start_row: 4, start_col: 0, end_row: 4, end_col: 1 },
    ];
    bands::insert(&mut book, "Sheet1", Band::Rows, 2, 1);
    let merges = &book.sheets[0].merge_cells;
    assert_eq!((merges[0].start_row, merges[0].end_row), (1, 4), "grew");
    assert_eq!((merges[1].start_row, merges[1].end_row), (5, 5), "slid");
}

#[test]
fn a_merge_whose_every_row_went_goes_with_them() {
    let mut book = a_workbook();
    book.sheets[0].merge_cells = vec![
        MergeCell { start_row: 2, start_col: 0, end_row: 3, end_col: 0 },
        MergeCell { start_row: 1, start_col: 0, end_row: 4, end_col: 0 },
    ];
    bands::remove(&mut book, "Sheet1", Band::Rows, 2, 2);
    let merges = &book.sheets[0].merge_cells;
    assert_eq!(merges.len(), 1, "the one wholly inside the band is gone");
    assert_eq!(
        (merges[0].start_row, merges[0].end_row),
        (1, 2),
        "and the one that only overlapped shrank",
    );
}

#[test]
fn a_freeze_above_the_insert_moves_and_one_below_it_does_not() {
    // Holding the first two rows in view: a row put in ABOVE the fold is held
    // too, so the fold comes down. A row put in below it is not.
    let mut above = a_workbook();
    above.sheets[0].frozen_rows = 2;
    bands::insert(&mut above, "Sheet1", Band::Rows, 1, 1);
    assert_eq!(above.sheets[0].frozen_rows, 3);

    let mut below = a_workbook();
    below.sheets[0].frozen_rows = 2;
    bands::insert(&mut below, "Sheet1", Band::Rows, 3, 1);
    assert_eq!(below.sheets[0].frozen_rows, 2);

    // A row put in AT the fold is held too.
    let mut at_the_fold = a_workbook();
    at_the_fold.sheets[0].frozen_rows = 2;
    bands::insert(&mut at_the_fold, "Sheet1", Band::Rows, 2, 1);
    assert_eq!(at_the_fold.sheets[0].frozen_rows, 3);
}

#[test]
fn a_freeze_loses_only_the_rows_it_was_holding() {
    // Two rows frozen, rows 2 to 4 taken out: only ONE of the three was ever
    // held, so one is left. Subtracting the whole band and clamping at zero
    // would unfreeze the sheet.
    let mut book = a_workbook();
    book.sheets[0].frozen_rows = 2;
    bands::remove(&mut book, "Sheet1", Band::Rows, 2, 3);
    assert_eq!(book.sheets[0].frozen_rows, 1);

    // Taking out rows below the fold leaves it alone.
    let mut below = a_workbook();
    below.sheets[0].frozen_rows = 2;
    bands::remove(&mut below, "Sheet1", Band::Rows, 3, 1);
    assert_eq!(below.sheets[0].frozen_rows, 2);

    // And taking out everything it held unfreezes it.
    let mut all = a_workbook();
    all.sheets[0].frozen_rows = 2;
    bands::remove(&mut all, "Sheet1", Band::Rows, 1, 2);
    assert_eq!(all.sheets[0].frozen_rows, 0);
}

#[test]
fn a_table_the_row_lands_inside_grows() {
    let mut book = a_workbook();
    book.sheets[0].tables = vec![Table {
        name: "Staff".to_string(),
        columns: vec!["ID".to_string(), "PAY".to_string()],
        start_row: 1,
        start_col: 0,
        end_row: 3,
        end_col: 1,
        header_rows: 1,
        style: None,
        banded_rows: false,
        accent: None,
        band: None,
        rule: None,
        outline: None,
    }];
    bands::insert(&mut book, "Sheet1", Band::Rows, 2, 1);
    let table = &book.sheets[0].tables[0];
    assert_eq!((table.start_row, table.end_row), (1, 4));
}

#[test]
fn an_inserted_column_takes_a_width_and_pushes_the_others_along() {
    let mut book = a_workbook();
    bands::insert(&mut book, "Sheet1", Band::Columns, 1, 1);
    assert_eq!(number_at(&book, "Sheet1", 1, 0), Some(1.0), "A stayed");
    assert_eq!(number_at(&book, "Sheet1", 1, 1), None, "B is now empty");
    assert_eq!(
        formula_at(&book, "Sheet1", 1, 2).as_deref(),
        Some("=A1*10"),
        "the formula moved to C and still reads A",
    );
    let widths = &book.sheets[0].col_widths;
    assert_eq!(widths[0], 8.0, "A kept its width");
    assert_eq!(widths[1], 8.43, "the new column wears the default");
    assert_eq!(widths[2], 9.0, "and B's width went with B");
}

#[test]
fn a_removed_column_takes_its_cells_and_its_width() {
    let mut book = a_workbook();
    bands::remove(&mut book, "Sheet1", Band::Columns, 0, 1);
    assert_eq!(
        formula_at(&book, "Sheet1", 1, 0).as_deref(),
        Some("=#REF!*10"),
        "the formula came left, and what it read is gone",
    );
    assert_eq!(book.sheets[0].col_widths[0], 9.0, "B's width, now A's");
}

#[test]
fn putting_no_rows_in_changes_nothing() {
    let mut book = a_workbook();
    let before = format!("{:?}", book.sheets[0].rows);
    bands::insert(&mut book, "Sheet1", Band::Rows, 2, 0);
    assert_eq!(format!("{:?}", book.sheets[0].rows), before);
}

#[test]
fn a_sheet_that_is_not_there_is_left_alone() {
    let mut book = a_workbook();
    let before = format!("{:?}", book.sheets[0].rows);
    bands::insert(&mut book, "Nowhere", Band::Rows, 2, 1);
    assert_eq!(format!("{:?}", book.sheets[0].rows), before);
}
