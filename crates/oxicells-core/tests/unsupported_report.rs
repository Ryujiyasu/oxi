// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! What a sheet says about the parts of itself this build cannot show.
//!
//! The fixture was written by Excel 16 and carries a conditional format, a
//! data validation and a hyperlink — none of which reach the IR.

use oxicells_core::parser::parse_xlsx;

const PLAIN: &[u8] = include_bytes!("fixtures/hidden_rows_cols.xlsx");
const RICH: &[u8] = include_bytes!("fixtures/unsupported_bits.xlsx");

#[test]
fn a_sheet_names_what_it_could_not_show() {
    let workbook = parse_xlsx(RICH).expect("the fixture parses");
    let mut noted = workbook.sheets[0].unsupported_elements.clone();
    noted.sort();
    assert_eq!(
        noted,
        vec![
            "Conditional formatting".to_string(),
            "Data validation".to_string(),
            "Hyperlinks".to_string(),
        ]
    );
}

#[test]
fn a_plain_sheet_names_nothing() {
    let workbook = parse_xlsx(PLAIN).expect("the fixture parses");
    assert!(workbook.sheets[0].unsupported_elements.is_empty());
}
