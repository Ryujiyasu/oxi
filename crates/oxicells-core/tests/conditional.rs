// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! Which cells a conditional rule catches, and what it puts on them.
//!
//! The 285-workbook corpus cannot answer this: 3288 of its 3342 rule ranges
//! look for a tick mark on a blank government form, so the right answer for
//! almost all of them is "nothing is formatted" — a test that passes by doing
//! nothing. These workbooks are written by `tools/metrics/_cf_repro_gen.py`
//! instead, one rule kind at a time with values chosen so that some cells match
//! and some do not, and every expectation below was first read off Excel
//! through `Range.DisplayFormat` by
//! `tools/metrics/xlsx_conditional_agreement.py` rather than reasoned out.
//!
//! Rows and columns are counted from zero throughout, the way the rules state
//! their own ranges.

use oxicells_core::conditional::{self, Caught};

const CELL_IS: &[u8] = include_bytes!("../../../tests/fixtures/conditional/cell_is.xlsx");
const TEXT: &[u8] = include_bytes!("../../../tests/fixtures/conditional/text_rules.xlsx");
const BLANKS: &[u8] = include_bytes!("../../../tests/fixtures/conditional/blanks.xlsx");
const DUPLICATES: &[u8] = include_bytes!("../../../tests/fixtures/conditional/duplicates.xlsx");
const EXPRESSION: &[u8] = include_bytes!("../../../tests/fixtures/conditional/expression.xlsx");
const TOP10: &[u8] = include_bytes!("../../../tests/fixtures/conditional/top10.xlsx");
const LAYERED: &[u8] = include_bytes!("../../../tests/fixtures/conditional/layered.xlsx");
const STOPPING: &[u8] = include_bytes!("../../../tests/fixtures/conditional/stop_if_true.xlsx");

const RED: &str = "FFC7CE";
const GREEN: &str = "C6EFCE";
const AMBER: &str = "FFEB9C";
const BLUE: &str = "BDD7EE";

fn hits(bytes: &[u8]) -> Vec<Caught> {
    let book = oxicells_core::parse_xlsx(bytes).expect("the fixture parses");
    conditional::caught(&book, 0)
}

/// Every cell wearing one fill, as `(row, column)` — the fills are far enough
/// apart that a wrong rule cannot borrow a right colour.
fn wearing(bytes: &[u8], fill: &str) -> Vec<(u32, u32)> {
    let mut found: Vec<(u32, u32)> = hits(bytes)
        .into_iter()
        .filter(|hit| hit.style.bg_color.as_deref() == Some(fill))
        .map(|hit| (hit.row, hit.col))
        .collect();
    found.sort_unstable();
    found
}

#[test]
fn a_number_against_a_bound() {
    // A2:A8 holds 10, 50, 75, 100, 3, nothing, and the TEXT "50".
    // greaterThan 50 takes 75 and 100 — and the text, because Excel orders
    // every piece of text above every number rather than reading fifty out of
    // it. Not 50 itself, and not the empty cell.
    assert_eq!(wearing(CELL_IS, RED), vec![(3, 0), (4, 0), (7, 0)]);
    // between 5 and 20 takes 10 alone; 3 is under it and 50 over.
    assert_eq!(wearing(CELL_IS, GREEN), vec![(1, 0)]);
}

#[test]
fn text_read_without_regard_to_case() {
    // apple, Pineapple, banana, APPLESAUCE, nothing, cape.
    assert_eq!(wearing(TEXT, RED), vec![(1, 0), (2, 0), (4, 0)]);
    assert_eq!(wearing(TEXT, GREEN), vec![(3, 0)]);
    assert_eq!(wearing(TEXT, AMBER), vec![(6, 0)]);
}

#[test]
fn blank_and_not_blank_divide_the_range() {
    let empty = wearing(BLANKS, RED);
    let full = wearing(BLANKS, GREEN);
    // A4 holds the empty string and A6 holds nothing at all — the file records
    // no cell for it, and Excel colours it just the same.
    assert_eq!(empty, vec![(2, 0), (3, 0), (5, 0)]);
    assert_eq!(full, vec![(1, 0), (4, 0)]);
    // Between them they account for every cell in A2:A6 exactly once: a rule
    // that catches nothing and a rule that catches everything would both pass a
    // test that only looked at one of them.
    let mut both: Vec<_> = empty.into_iter().chain(full).collect();
    both.sort_unstable();
    assert_eq!(both, vec![(1, 0), (2, 0), (3, 0), (4, 0), (5, 0)]);
}

#[test]
fn counting_rules_read_the_whole_range() {
    // red, blue, red, green, blue, red.
    assert_eq!(
        wearing(DUPLICATES, RED),
        vec![(1, 0), (2, 0), (3, 0), (5, 0), (6, 0)]
    );
    assert_eq!(wearing(DUPLICATES, GREEN), vec![(4, 0)]);
}

#[test]
fn a_formula_is_read_at_the_cell_it_stands_in() {
    // A2:A5 is 4, 11, 40, 9 against a fixed 10 in B2.
    assert_eq!(wearing(EXPRESSION, RED), vec![(2, 0), (3, 0)]);
    // B2:B5 carries `MOD(ROW(),2)=0` — banded rows, the commonest conditional
    // rule there is. `ROW()` with no argument only has an answer if the formula
    // is read AT a cell rather than merely about one, and read nowhere it
    // returns #VALUE! and the rule silently catches nothing.
    assert_eq!(wearing(EXPRESSION, AMBER), vec![(1, 1), (3, 1)]);
}

#[test]
fn the_rank_rules() {
    // A2:A7 is 12, 55, 3, 98, 40, 77.
    assert_eq!(wearing(TOP10, RED), vec![(4, 0), (6, 0)]);
    assert_eq!(wearing(TOP10, GREEN), vec![(1, 0), (3, 0)]);
    // B2:B7 is 5, 90, 40, 70, 20, 60, and the rule asks for the top 30 per
    // cent. Excel takes one cell of the six, not two: the share is rounded
    // down, which is measured rather than assumed.
    assert_eq!(wearing(TOP10, AMBER), vec![(2, 1)]);
}

#[test]
fn two_rules_over_one_cell_are_laid_over_each_other() {
    // A2:A7 is 1..6. The rule with the higher precedence sets weight and no
    // fill; the one under it sets fill and no weight. Excel shows both, so
    // picking a winning rule loses one of them.
    let found = hits(LAYERED);
    let heavy: Vec<(u32, u32)> = found
        .iter()
        .filter(|hit| hit.style.bold == Some(true))
        .map(|hit| (hit.row, hit.col))
        .collect();
    assert_eq!(heavy, vec![(4, 0), (5, 0), (6, 0)]);
    assert_eq!(
        wearing(LAYERED, BLUE),
        vec![(2, 0), (3, 0), (4, 0), (5, 0), (6, 0)]
    );
    // And the cells that got both really did get both, in one entry.
    let together = found
        .iter()
        .find(|hit| (hit.row, hit.col) == (5, 0))
        .expect("A6 is caught");
    assert_eq!(together.style.bold, Some(true));
    assert_eq!(together.style.bg_color.as_deref(), Some(BLUE));
}

#[test]
fn a_rule_that_says_stop_ends_the_stack() {
    // The first rule takes 5 and 6 and stops; without honouring that, the rule
    // under it would repaint them.
    assert_eq!(wearing(STOPPING, RED), vec![(5, 0), (6, 0)]);
    assert_eq!(wearing(STOPPING, GREEN), vec![(2, 0), (3, 0), (4, 0)]);
}

#[test]
fn a_cell_is_reported_once() {
    for book in [CELL_IS, TEXT, BLANKS, DUPLICATES, EXPRESSION, TOP10, LAYERED, STOPPING] {
        let found = hits(book);
        let mut places: Vec<(u32, u32)> = found.iter().map(|hit| (hit.row, hit.col)).collect();
        let before = places.len();
        places.sort_unstable();
        places.dedup();
        assert_eq!(places.len(), before, "a cell was reported twice");
    }
}
