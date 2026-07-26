// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! A1-style references: `A1`, `$A$1`, `A1:B2`, `Sheet1!A1`, `'My Sheet'!A1`.
//!
//! Rows and columns are both **0-based** inside this crate. The 1-based row
//! numbers and letter columns exist only at the parse and format boundaries.

use std::fmt;

/// Highest column index Excel allows (`XFD`), 0-based.
pub const MAX_COL: u32 = 16_383;
/// Highest row index Excel allows (row 1_048_576), 0-based.
pub const MAX_ROW: u32 = 1_048_575;

/// A single cell reference, carrying whether each half was written absolute.
///
/// The `$` markers do not affect evaluation, but they must survive a parse and
/// re-format round trip so that a formula can be rewritten without silently
/// changing what happens when a user later copies the cell.
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash)]
pub struct CellRef {
    pub col: u32,
    pub row: u32,
    pub col_absolute: bool,
    pub row_absolute: bool,
}

impl CellRef {
    pub fn new(col: u32, row: u32) -> CellRef {
        CellRef {
            col,
            row,
            col_absolute: false,
            row_absolute: false,
        }
    }

    /// The (col, row) pair, discarding absoluteness. This is the identity used
    /// as a key in the dependency graph.
    pub fn coord(self) -> (u32, u32) {
        (self.col, self.row)
    }

    pub fn to_a1(self) -> String {
        format!(
            "{}{}{}{}",
            if self.col_absolute { "$" } else { "" },
            col_to_letters(self.col),
            if self.row_absolute { "$" } else { "" },
            self.row + 1
        )
    }
}

impl fmt::Display for CellRef {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.write_str(&self.to_a1())
    }
}

/// A rectangular range. A single cell is a range whose corners are equal.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct RangeRef {
    pub start: CellRef,
    pub end: CellRef,
}

impl RangeRef {
    pub fn single(cell: CellRef) -> RangeRef {
        RangeRef {
            start: cell,
            end: cell,
        }
    }

    /// Build a range from two corners written in any order.
    ///
    /// `=SUM(B2:A1)` is legal in Excel and means the same as `=SUM(A1:B2)`, so
    /// the corners are normalised here rather than at every use site.
    pub fn normalised(a: CellRef, b: CellRef) -> RangeRef {
        let mut start = a;
        let mut end = b;
        if start.col > end.col {
            std::mem::swap(&mut start.col, &mut end.col);
            std::mem::swap(&mut start.col_absolute, &mut end.col_absolute);
        }
        if start.row > end.row {
            std::mem::swap(&mut start.row, &mut end.row);
            std::mem::swap(&mut start.row_absolute, &mut end.row_absolute);
        }
        RangeRef { start, end }
    }

    pub fn is_single(&self) -> bool {
        self.start.coord() == self.end.coord()
    }

    pub fn contains(&self, col: u32, row: u32) -> bool {
        col >= self.start.col && col <= self.end.col && row >= self.start.row && row <= self.end.row
    }

    pub fn width(&self) -> u32 {
        self.end.col - self.start.col + 1
    }

    pub fn height(&self) -> u32 {
        self.end.row - self.start.row + 1
    }

    /// Iterate the range in row-major order, which is the order Excel's
    /// range enumeration and `INDEX` addressing use.
    pub fn iter(&self) -> impl Iterator<Item = (u32, u32)> + '_ {
        let (c0, c1) = (self.start.col, self.end.col);
        (self.start.row..=self.end.row).flat_map(move |r| (c0..=c1).map(move |c| (c, r)))
    }

    /// Intersection of two ranges, or `None` when they do not overlap
    /// (Excel's `#NULL!` case).
    pub fn intersect(&self, other: &RangeRef) -> Option<RangeRef> {
        let col0 = self.start.col.max(other.start.col);
        let col1 = self.end.col.min(other.end.col);
        let row0 = self.start.row.max(other.start.row);
        let row1 = self.end.row.min(other.end.row);
        if col0 > col1 || row0 > row1 {
            return None;
        }
        Some(RangeRef {
            start: CellRef::new(col0, row0),
            end: CellRef::new(col1, row1),
        })
    }
}

impl fmt::Display for RangeRef {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        if self.is_single() {
            f.write_str(&self.start.to_a1())
        } else {
            write!(f, "{}:{}", self.start.to_a1(), self.end.to_a1())
        }
    }
}

/// A range together with the sheet it lives on. `sheet: None` means "the sheet
/// the formula is on".
#[derive(Debug, Clone, PartialEq)]
pub struct Reference {
    pub sheet: Option<String>,
    pub range: RangeRef,
}

impl fmt::Display for Reference {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match &self.sheet {
            Some(name) if needs_quoting(name) => write!(f, "'{}'!{}", name.replace('\'', "''"), self.range),
            Some(name) => write!(f, "{}!{}", name, self.range),
            None => write!(f, "{}", self.range),
        }
    }
}

fn needs_quoting(name: &str) -> bool {
    name.is_empty()
        || name
            .chars()
            .any(|c| !(c.is_alphanumeric() || c == '_' || c == '.'))
        || name.chars().next().is_some_and(|c| c.is_ascii_digit())
}

/// Convert a 0-based column index to letters: 0 → `A`, 26 → `AA`.
pub fn col_to_letters(mut col: u32) -> String {
    let mut out = Vec::new();
    loop {
        out.push(b'A' + (col % 26) as u8);
        if col < 26 {
            break;
        }
        // Excel's column names are bijective base-26, not plain base-26:
        // there is no zero digit, so the carry subtracts one.
        col = col / 26 - 1;
    }
    out.reverse();
    String::from_utf8(out).expect("ASCII only")
}

/// Convert column letters to a 0-based index. Returns `None` when the letters
/// are empty, non-alphabetic, or address a column past `XFD`.
pub fn letters_to_col(s: &str) -> Option<u32> {
    if s.is_empty() || s.len() > 3 {
        return None;
    }
    let mut col: u32 = 0;
    for ch in s.chars() {
        let d = match ch {
            'A'..='Z' => ch as u32 - 'A' as u32,
            'a'..='z' => ch as u32 - 'a' as u32,
            _ => return None,
        };
        col = col.checked_mul(26)?.checked_add(d + 1)?;
    }
    let col = col - 1;
    (col <= MAX_COL).then_some(col)
}

/// Parse a single A1-style cell reference, with or without `$` markers.
///
/// Returns `None` for anything that is not a well-formed in-range reference,
/// which is how the parser distinguishes `A1` from a defined name like `TAX`
/// or a function name like `LOG10`.
pub fn parse_a1(s: &str) -> Option<CellRef> {
    let bytes = s.as_bytes();
    let mut i = 0;

    let col_absolute = bytes.first() == Some(&b'$');
    if col_absolute {
        i += 1;
    }

    let letters_start = i;
    while i < bytes.len() && bytes[i].is_ascii_alphabetic() {
        i += 1;
    }
    if i == letters_start {
        return None;
    }
    let col = letters_to_col(&s[letters_start..i])?;

    let row_absolute = bytes.get(i) == Some(&b'$');
    if row_absolute {
        i += 1;
    }

    let digits_start = i;
    while i < bytes.len() && bytes[i].is_ascii_digit() {
        i += 1;
    }
    if i == digits_start || i != bytes.len() {
        return None;
    }
    let row_1based: u32 = s[digits_start..i].parse().ok()?;
    if row_1based == 0 {
        return None;
    }
    let row = row_1based - 1;
    if row > MAX_ROW {
        return None;
    }

    Some(CellRef {
        col,
        row,
        col_absolute,
        row_absolute,
    })
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn column_names_are_bijective_base26() {
        assert_eq!(col_to_letters(0), "A");
        assert_eq!(col_to_letters(25), "Z");
        assert_eq!(col_to_letters(26), "AA");
        assert_eq!(col_to_letters(701), "ZZ");
        assert_eq!(col_to_letters(702), "AAA");
        assert_eq!(col_to_letters(MAX_COL), "XFD");
    }

    #[test]
    fn column_letters_round_trip() {
        for col in [0u32, 1, 25, 26, 27, 701, 702, 16_382, MAX_COL] {
            assert_eq!(letters_to_col(&col_to_letters(col)), Some(col), "col {col}");
        }
    }

    #[test]
    fn columns_past_xfd_are_rejected() {
        assert_eq!(letters_to_col("XFD"), Some(MAX_COL));
        assert_eq!(letters_to_col("XFE"), None);
        assert_eq!(letters_to_col("AAAA"), None);
    }

    #[test]
    fn absolute_markers_survive_a_round_trip() {
        for text in ["A1", "$A1", "A$1", "$A$1", "$XFD$1048576"] {
            let parsed = parse_a1(text).unwrap_or_else(|| panic!("failed to parse {text}"));
            assert_eq!(parsed.to_a1(), text);
        }
    }

    #[test]
    fn non_references_are_rejected() {
        // These must fall through to the function-name / defined-name paths.
        for text in ["A", "1", "A0", "A1B", "", "$", "SUM", "TAX_RATE", "A1048577"] {
            assert_eq!(parse_a1(text), None, "should not parse {text:?}");
        }
    }

    #[test]
    fn function_looking_names_can_be_real_addresses() {
        // `LOG10` is a genuine cell address: column LOG, row 10. Excel treats
        // `=LOG10` as that cell and `=LOG10(100)` as the function, so the
        // ambiguity cannot be resolved here — only by looking for a following
        // `(` in the parser. This is why the lexer does not classify names.
        assert_eq!(parse_a1("LOG10"), Some(CellRef::new(8508, 9)));
        assert_eq!(col_to_letters(8508), "LOG");
    }

    #[test]
    fn ranges_normalise_reversed_corners() {
        let a = parse_a1("B2").unwrap();
        let b = parse_a1("A1").unwrap();
        let range = RangeRef::normalised(a, b);
        assert_eq!(range.to_string(), "A1:B2");
        assert_eq!(range.width(), 2);
        assert_eq!(range.height(), 2);
    }

    #[test]
    fn ranges_iterate_in_row_major_order() {
        let range = RangeRef::normalised(parse_a1("A1").unwrap(), parse_a1("B2").unwrap());
        let visited: Vec<_> = range.iter().collect();
        assert_eq!(visited, vec![(0, 0), (1, 0), (0, 1), (1, 1)]);
    }

    #[test]
    fn disjoint_ranges_do_not_intersect() {
        let a = RangeRef::normalised(parse_a1("A1").unwrap(), parse_a1("B2").unwrap());
        let b = RangeRef::normalised(parse_a1("D4").unwrap(), parse_a1("E5").unwrap());
        assert_eq!(a.intersect(&b), None);
        assert!(a.intersect(&a).is_some());
    }

    #[test]
    fn sheet_names_are_quoted_only_when_required() {
        let range = RangeRef::single(parse_a1("A1").unwrap());
        let plain = Reference {
            sheet: Some("Sheet1".into()),
            range,
        };
        let spaced = Reference {
            sheet: Some("My Sheet".into()),
            range,
        };
        assert_eq!(plain.to_string(), "Sheet1!A1");
        assert_eq!(spaced.to_string(), "'My Sheet'!A1");
    }
}
