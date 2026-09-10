// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! What a sheet says about the parts of itself this build cannot show.
//!
//! The fixture was written by Excel 16 and carries a conditional format, a
//! data validation and a hyperlink. All three now reach the IR, so the sheet
//! names nothing: a report of what cannot be shown is worth less than nothing
//! when it lists what can. What each of them turned into is asserted below, so
//! an empty report stays a claim about the parser and not about the report.

use oxicells_core::parser::parse_xlsx;

const PLAIN: &[u8] = include_bytes!("fixtures/hidden_rows_cols.xlsx");
const RICH: &[u8] = include_bytes!("fixtures/unsupported_bits.xlsx");

#[test]
fn a_sheet_names_what_it_could_not_show() {
    let workbook = parse_xlsx(RICH).expect("the fixture parses");
    let mut noted = workbook.sheets[0].unsupported_elements.clone();
    noted.sort();
    assert!(noted.is_empty(), "nothing should be outstanding: {noted:?}");
}

/// The conditional format reaches the IR with the range it covers and the look
/// it puts on, which is why the report above names nothing.
#[test]
fn a_sheet_carries_the_rule_that_redresses_it() {
    let workbook = parse_xlsx(RICH).expect("the fixture parses");
    let rules = &workbook.sheets[0].conditional_rules;
    assert_eq!(rules.len(), 1, "one rule was expected: {rules:?}");
    assert!(!rules[0].ranges.is_empty(), "the rule covers something");
    let style = rules[0].style.as_ref().expect("the rule names a dxf");
    assert!(
        style.font_color.is_some() || style.bg_color.is_some() || style.bold.is_some(),
        "the dxf changes something: {style:?}"
    );
}

/// The validation the same fixture holds reaches the IR with its range and the
/// rule it applies, which is why the report above no longer names it.
#[test]
fn a_sheet_carries_the_rule_its_range_is_under() {
    let workbook = parse_xlsx(RICH).expect("the fixture parses");
    let rules = &workbook.sheets[0].validations;
    assert_eq!(rules.len(), 1, "one rule was expected: {rules:?}");
    assert!(!rules[0].kind.is_empty(), "the rule says what it tests");
    assert!(!rules[0].ranges.is_empty(), "the rule covers something");
}

/// The link the same fixture holds reaches the IR with somewhere to go, which
/// is why the report above no longer names it.
#[test]
fn a_sheet_carries_the_link_its_cell_holds() {
    let workbook = parse_xlsx(RICH).expect("the fixture parses");
    let links = &workbook.sheets[0].hyperlinks;
    assert_eq!(links.len(), 1, "one link was expected: {links:?}");
    assert!(!links[0].target.is_empty(), "the link goes nowhere");
}

#[test]
fn a_plain_sheet_names_nothing() {
    let workbook = parse_xlsx(PLAIN).expect("the fixture parses");
    assert!(workbook.sheets[0].unsupported_elements.is_empty());
}
