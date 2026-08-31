// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! Putting rows and columns in, and taking them out.
//!
//! Among the most ordinary things anyone does to a spreadsheet, and the one
//! that touches the most of it. A row does not move on its own: the formulas
//! that mention it move, on every sheet and not only its own; the merges over
//! it stretch or slide; the table it falls inside grows; the frozen rows above
//! it stay put while the ones below come down; the pictures hanging off it
//! follow.
//!
//! Getting any one of those wrong is the kind of mistake nobody notices until
//! the numbers are already wrong, so each is done deliberately here rather than
//! left to happen.
//!
//! The hard part — rewriting the formulas — is not done here. `oxicells-calc`
//! already knows how, including that a reference to something taken away
//! becomes `#REF!` and that a range an insertion lands inside grows to hold it.

use crate::ir::{Anchor, Sheet, Workbook};
use oxicells_calc::{shift_formula_references, ReferenceShift, ShiftAxis};

/// Which way a band of rows or columns runs.
pub use oxicells_calc::ShiftAxis as Band;

/// Put `count` rows or columns into `sheet` at `at`, pushing what was there
/// along.
///
/// `at` counts from one for rows and from zero for columns, which is how the
/// IR counts each of them elsewhere. Everything from `at` onwards moves by
/// `count`; a merge, table or picture that straddles `at` grows instead.
pub fn insert(workbook: &mut Workbook, sheet: &str, band: Band, at: u32, count: u32) {
    if count > 0 {
        move_band(workbook, sheet, band, at, count as i64);
    }
}

/// Take `count` rows or columns out of `sheet` at `at`, closing the gap.
///
/// A formula reading anything inside the band becomes `#REF!`, as Excel's
/// does: the cells it named are gone, and shifting the reference along would
/// quietly answer with somebody else's number.
pub fn remove(workbook: &mut Workbook, sheet: &str, band: Band, at: u32, count: u32) {
    if count > 0 {
        move_band(workbook, sheet, band, at, -(count as i64));
    }
}

fn move_band(workbook: &mut Workbook, sheet: &str, band: Band, at: u32, count: i64) {
    // The formulas first, and across the WHOLE workbook: `=Data!A5` on another
    // sheet follows a row inserted on Data.
    reword_formulas(workbook, sheet, band, at, count);
    reword_names(workbook, sheet, band, at, count);

    let Some(target) = workbook.sheets.iter_mut().find(|one| one.name == sheet) else {
        return;
    };
    match band {
        Band::Rows => move_rows(target, at, count),
        Band::Columns => move_columns(target, at, count),
    }
    move_merges(target, band, at, count);
    move_tables(target, band, at, count);
    move_panes(target, band, at, count);
    move_filter(target, band, at, count);
    move_hangers(target, band, at, count);
    move_declared_range(target, band, at, count);
}

// ── the formulas ────────────────────────────────────────────────────────────

fn reword_formulas(workbook: &mut Workbook, sheet: &str, band: Band, at: u32, count: i64) {
    // `ReferenceShift` counts both axes from one, and reaches across the whole
    // of the other one: a row goes right across the sheet.
    let axis = match band {
        Band::Rows => ShiftAxis::Rows,
        Band::Columns => ShiftAxis::Columns,
    };
    let at = match band {
        Band::Rows => at,
        Band::Columns => at + 1,
    };
    for other in &mut workbook.sheets {
        // Which sheet these formulas are written on decides what an
        // unqualified `A1` means. Without saying so, a row put into one sheet
        // would drag every other sheet's references along with it.
        let shift = ReferenceShift {
            axis,
            at,
            count,
            across: (1, u32::MAX),
            sheet: Some(sheet),
            on_sheet: Some(other.name.as_str()),
        };
        for row in &mut other.rows {
            for cell in &mut row.cells {
                let Some(formula) = cell.formula.as_deref() else {
                    continue;
                };
                // A formula this cannot read is left exactly as it was. Half a
                // rewrite is worse than none: it would move some references
                // and not others.
                if let Ok(moved) = shift_formula_references(formula, &shift) {
                    cell.formula = Some(moved);
                }
            }
        }
    }
}

/// A name points at cells, so it moves with them.
///
/// Asked of Excel, a name standing for `Sheet1!$B$3:$B$5`: a row put in above
/// leaves `$B$4:$B$6`, one put in the middle of it leaves `$B$3:$B$6`, taking
/// the middle row out leaves `$B$3:$B$4`, and taking all three out leaves
/// `#REF!`. A name reaching PAST a part-width band is left alone, exactly as a
/// formula is — pushing `B2` down moves a name on `$B$3:$B$5` but not one on
/// `$A$3:$C$5`.
///
/// A name always says which sheet it means, so there is no home sheet to read
/// an unqualified reference against.
fn reword_names(workbook: &mut Workbook, sheet: &str, band: Band, at: u32, count: i64) {
    let axis = match band {
        Band::Rows => ShiftAxis::Rows,
        Band::Columns => ShiftAxis::Columns,
    };
    let at = match band {
        Band::Rows => at,
        Band::Columns => at + 1,
    };
    let shift = ReferenceShift {
        axis,
        at,
        count,
        across: (1, u32::MAX),
        sheet: Some(sheet),
        on_sheet: None,
    };
    for (_, refers_to) in &mut workbook.defined_names {
        // One this crate cannot read keeps the text it had, the same as a
        // formula it cannot read.
        if let Ok(moved) = shift_formula_references(refers_to, &shift) {
            *refers_to = moved;
        }
    }
}

// ── the cells ───────────────────────────────────────────────────────────────

fn move_rows(sheet: &mut Sheet, at: u32, count: i64) {
    if count < 0 {
        let gone = (-count) as u32;
        sheet
            .rows
            .retain(|row| row.index < at || row.index >= at + gone);
    }
    for row in &mut sheet.rows {
        if row.index >= at {
            // A row's own height and hidden flag travel with it, being part of
            // the row rather than of the position.
            row.index = shifted_one(row.index, at, count);
        }
    }
    sheet.rows.sort_by_key(|row| row.index);
}

fn move_columns(sheet: &mut Sheet, at: u32, count: i64) {
    for row in &mut sheet.rows {
        if count < 0 {
            let gone = (-count) as u32;
            row.cells
                .retain(|cell| cell.col < at || cell.col >= at + gone);
        }
        for cell in &mut row.cells {
            if cell.col >= at {
                cell.col = shifted_zero(cell.col, at, count);
            }
        }
        row.cells.sort_by_key(|cell| cell.col);
    }

    // A column's width belongs to the column, so it moves with it.
    let place = (at as usize).min(sheet.col_widths.len());
    if count > 0 {
        let width = sheet.default_col_width;
        for _ in 0..count {
            sheet.col_widths.insert(place, width);
        }
        sheet.col_count += count as usize;
    } else {
        let gone = ((-count) as usize).min(sheet.col_widths.len().saturating_sub(place));
        sheet.col_widths.drain(place..place + gone);
        sheet.col_count = sheet.col_count.saturating_sub((-count) as usize);
    }

    sheet.hidden_cols = sheet
        .hidden_cols
        .iter()
        .filter(|col| count > 0 || **col < at || **col >= at + (-count) as u32)
        .map(|col| if *col >= at { shifted_zero(*col, at, count) } else { *col })
        .collect();

    for (first, last, ..) in &mut sheet.col_fonts {
        let (moved_first, moved_last) = both_ends_zero(*first, *last, at, count);
        *first = moved_first;
        *last = moved_last;
    }
}

// ── everything drawn over or around them ────────────────────────────────────

fn move_merges(sheet: &mut Sheet, band: Band, at: u32, count: i64) {
    sheet.merge_cells.retain_mut(|merge| {
        let (first, last) = match band {
            Band::Rows => (&mut merge.start_row, &mut merge.end_row),
            Band::Columns => (&mut merge.start_col, &mut merge.end_col),
        };
        let one_based = matches!(band, Band::Rows);
        let Some((moved_first, moved_last)) = span(*first, *last, at, count, one_based) else {
            // Every row it covered is gone, so the merge is gone.
            return false;
        };
        *first = moved_first;
        *last = moved_last;
        // A merge of one cell is not a merge.
        moved_last > moved_first || {
            let (other_first, other_last) = match band {
                Band::Rows => (merge.start_col, merge.end_col),
                Band::Columns => (merge.start_row, merge.end_row),
            };
            other_last > other_first
        }
    });
}

fn move_tables(sheet: &mut Sheet, band: Band, at: u32, count: i64) {
    sheet.tables.retain_mut(|table| {
        let (first, last) = match band {
            Band::Rows => (&mut table.start_row, &mut table.end_row),
            Band::Columns => (&mut table.start_col, &mut table.end_col),
        };
        let one_based = matches!(band, Band::Rows);
        let Some((moved_first, moved_last)) = span(*first, *last, at, count, one_based) else {
            return false;
        };
        *first = moved_first;
        *last = moved_last;
        // A table whose columns went takes its headings with them.
        if matches!(band, Band::Columns) && count < 0 {
            let from = (at.saturating_sub(table.start_col)) as usize;
            let gone = ((-count) as usize).min(table.columns.len().saturating_sub(from));
            if from < table.columns.len() {
                table.columns.drain(from..from + gone);
            }
        }
        true
    });
}

fn move_panes(sheet: &mut Sheet, band: Band, at: u32, count: i64) {
    let held = match band {
        Band::Rows => &mut sheet.frozen_rows,
        Band::Columns => &mut sheet.frozen_cols,
    };
    if *held == 0 {
        return;
    }
    // A freeze holds a COUNT, not a position: `frozen_rows = 2` holds the
    // first two. In the sheet's own counting those are rows 1 and 2, or
    // columns 0 and 1, so the band starts one further along for rows.
    let first_held = match band {
        Band::Rows => 1,
        Band::Columns => 0,
    };
    let last_held = first_held + *held - 1;

    if count > 0 {
        // Rows put in AT or before the fold are held too, so the fold moves;
        // rows put in below it are not.
        if at <= last_held {
            *held += count as u32;
        }
        return;
    }
    // Taken out: what the freeze loses is the overlap with the band, which is
    // not the same as the band. Deleting rows 2 to 4 with two rows frozen
    // leaves one frozen, not none.
    let gone = (-count) as u32;
    let band_end = at + gone - 1;
    let overlap = last_held.min(band_end).saturating_sub(at.max(first_held)) + 1;
    if at <= last_held && band_end >= first_held {
        *held = held.saturating_sub(overlap.min(*held));
    }
}
fn move_filter(sheet: &mut Sheet, band: Band, at: u32, count: i64) {
    let Some(filter) = sheet.auto_filter.as_mut() else {
        return;
    };
    let (first, last) = match band {
        Band::Rows => (&mut filter.start_row, &mut filter.end_row),
        Band::Columns => (&mut filter.start_col, &mut filter.end_col),
    };
    match span(*first, *last, at, count, matches!(band, Band::Rows)) {
        Some((moved_first, moved_last)) => {
            *first = moved_first;
            *last = moved_last;
        }
        None => sheet.auto_filter = None,
    }
}

fn move_hangers(sheet: &mut Sheet, band: Band, at: u32, count: i64) {
    // An anchor counts both axes from zero, whichever way the sheet counts.
    let anchor_at = match band {
        Band::Rows => at.saturating_sub(1),
        Band::Columns => at,
    };
    let move_anchor = |anchor: &mut Anchor| {
        let held = match band {
            Band::Rows => &mut anchor.row,
            Band::Columns => &mut anchor.col,
        };
        if *held >= anchor_at {
            *held = shifted_zero(*held, anchor_at, count);
        }
    };
    for drawing in &mut sheet.drawings {
        move_anchor(&mut drawing.from);
        if let Some(to) = drawing.to.as_mut() {
            move_anchor(to);
        }
    }
    for comment in &mut sheet.comments {
        move_anchor(&mut comment.from);
        if let Some(to) = comment.to.as_mut() {
            move_anchor(to);
        }
    }
}

fn move_declared_range(sheet: &mut Sheet, band: Band, at: u32, count: i64) {
    let Some((first_row, first_col, last_row, last_col)) = sheet.declared_range else {
        return;
    };
    sheet.declared_range = match band {
        Band::Rows => span(first_row, last_row, at, count, true)
            .map(|(from, to)| (from, first_col, to, last_col)),
        Band::Columns => span(first_col, last_col, at, count, false)
            .map(|(from, to)| (first_row, from, last_row, to)),
    };
}

// ── the arithmetic, in one place ────────────────────────────────────────────

/// Where a single one-based index lands.
fn shifted_one(held: u32, at: u32, count: i64) -> u32 {
    if held < at {
        return held;
    }
    (held as i64 + count).max(at as i64) as u32
}

/// Where a single zero-based index lands.
fn shifted_zero(held: u32, at: u32, count: i64) -> u32 {
    if held < at {
        return held;
    }
    (held as i64 + count).max(at as i64) as u32
}

/// Where the two ends of a span land, or `None` when nothing of it is left.
///
/// A span the band lands inside GROWS — an inserted row inside a merge makes
/// the merge a row taller — while one the band only overlaps SHRINKS.
fn span(first: u32, last: u32, at: u32, count: i64, one_based: bool) -> Option<(u32, u32)> {
    let shifted = if one_based { shifted_one } else { shifted_zero };
    if count > 0 {
        let moved_first = shifted(first, at, count);
        // The far end moves when the insertion is at or before it, which is
        // what makes a span grow rather than slide when the band lands inside.
        let moved_last = if last >= at {
            (last as i64 + count) as u32
        } else {
            last
        };
        return Some((moved_first, moved_last));
    }
    let gone = (-count) as u32;
    let band_end = at + gone; // one past the last one taken
    if first >= at && last < band_end {
        return None; // all of it went
    }
    let take = |held: u32| -> u32 {
        if held < at {
            held
        } else if held < band_end {
            at // it was inside the band, so it collapses to where the band was
        } else {
            (held as i64 + count) as u32
        }
    };
    let (moved_first, moved_last) = (take(first), take(last));
    Some((moved_first, moved_last.max(moved_first)))
}

/// Both ends of a zero-based span, kept together.
fn both_ends_zero(first: u32, last: u32, at: u32, count: i64) -> (u32, u32) {
    span(first, last, at, count, false).unwrap_or((first, first))
}
