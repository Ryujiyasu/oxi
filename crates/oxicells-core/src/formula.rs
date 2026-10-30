// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! Formula evaluation, delegated to `oxicells-calc`.
//!
//! This module used to carry its own evaluator. It handled arithmetic and eleven
//! functions, had no comparison operators, and evaluated every formula against a
//! snapshot taken before evaluation began — so a formula that referenced another
//! formula read a stale value. Measured against the cached values in the 285
//! workbooks under `tools/golden-test/documents/xlsx`, it agreed with Excel on
//! about 15% of the 29,672 formula cells; `IF(A1>10,…)` alone produced `#VALUE!`
//! in over 23,000 of them.
//!
//! `oxicells-calc` builds a dependency graph and recalculates in topological
//! order. On the same corpus it agrees with Excel on 99.5%.
//!
//! # When to recalculate at all
//!
//! Usually: don't. An `.xlsx` already contains the values Excel computed, and
//! for display those values *are* the right answer — recalculating a freshly
//! loaded file can only introduce divergence. [`crate::parse_xlsx`] therefore
//! keeps them and calls [`fill_missing_formula_values`], which only computes
//! cells the file left without a cached value.
//!
//! Recalculate when the workbook has been *changed*: that is what
//! [`evaluate_workbook_formulas`] is for.

use crate::ir::{CellValue, Sheet, Workbook};
use oxicells_calc::reference::col_to_letters;
use oxicells_calc::{ExcelError, Value};

pub use oxicells_calc::{
    format_number, formula_from_r1c1, formula_to_r1c1, move_formula_references,
    rename_sheet_in_formula,
    shift_formula_references, translate_formula_references, transpose_formula_references, CellMove,
    ReferenceShift, ShiftAxis,
};

/// The workbook put to the engine once, for a caller with many expressions to
/// work out over the same book.
///
/// `evaluate_expression` assembles the book on every call, which is right for
/// the odd expression a macro evaluates and wrong for a conditional format,
/// which asks the same question of every cell in a range.
pub fn assemble_for_evaluation(
    sheets: &[Sheet],
    names: &[(String, String)],
    external: &[crate::ir::ExternalBook],
) -> oxicells_calc::Workbook {
    assemble_with_links(sheets, names, external, None)
}

/// Recalculate every formula in a single sheet, overwriting cached values.
///
/// Cross-sheet references cannot resolve here and become `#REF!`; use
/// [`evaluate_workbook_formulas`] when the workbook has more than one sheet.
/// How many of a workbook's formulas this engine cannot even read.
///
/// A formula that will not parse keeps the value the file was saved with, so
/// it costs nothing visible and agrees with any oracle by default. Counting
/// them is the only way a measurement can tell "we got this right" from "we
/// never looked".
pub fn unparsed_formulas(workbook: &Workbook) -> usize {
    workbook
        .sheets
        .iter()
        .flat_map(|sheet| sheet.rows.iter())
        .flat_map(|row| row.cells.iter())
        .filter_map(|cell| cell.formula.as_deref())
        .filter(|formula| {
            let text = formula.strip_prefix('=').unwrap_or(formula);
            oxicells_calc::parse(text).is_err()
        })
        .count()
}

pub fn evaluate_sheet_formulas(sheet: &mut Sheet) {
    // One sheet on its own carries no workbook, so it has no names either.
    recalculate(std::slice::from_mut(sheet), &[], &[], Overwrite::All, None, 1, true, &[]);
}

/// Recalculate every formula in the workbook, overwriting cached values.
///
/// Use after editing. Cross-sheet references resolve correctly.
pub fn evaluate_workbook_formulas(workbook: &mut Workbook) {
    let names = workbook.defined_names.clone();
    let external = workbook.external_books.clone();
    let report = recalculate(&mut workbook.sheets, &names, &external, Overwrite::All, None, workbook.calculation.compatibility_version, workbook.calculation.dbcs_language.unwrap_or(true), &workbook.recalculation.order);
    crate::recalculation::remember_completed_calculation(workbook, &report);
    if !report.deferred {
        workbook.recalculation.changed.clear();
        workbook.recalculation.dirty.clear();
        workbook.calculation.completed = None;
    }
}

/// The same, told what the moment is.
///
/// `now` is a serial: whole days since the last day of 1899, with the time of
/// day after the point. It is what `TODAY()` and `NOW()` answer.
///
/// Somewhere without a clock has to be told — a browser build has no
/// `SystemTime` — and somewhere with one may still want to pin the moment, so
/// that a sheet worked out twice comes out the same both times.
pub fn evaluate_workbook_formulas_at(workbook: &mut Workbook, now: f64) {
    let names = workbook.defined_names.clone();
    let external = workbook.external_books.clone();
    let report = recalculate(&mut workbook.sheets, &names, &external, Overwrite::All, Some(now), workbook.calculation.compatibility_version, workbook.calculation.dbcs_language.unwrap_or(true), &workbook.recalculation.order);
    crate::recalculation::remember_completed_calculation(workbook, &report);
    if !report.deferred {
        workbook.recalculation.changed.clear();
        workbook.recalculation.dirty.clear();
        workbook.calculation.completed = None;
    }
}

/// Recalculate using a temporary traversal order while retaining learned repairs
/// against the workbook's own persistent calculation chain.
pub fn evaluate_workbook_formulas_in_order(workbook: &mut Workbook, order: &[(String, (u32, u32))], now: Option<f64>) {
    let names = workbook.defined_names.clone();
    let external = workbook.external_books.clone();
    let report = recalculate(&mut workbook.sheets, &names, &external, Overwrite::All, now, workbook.calculation.compatibility_version, workbook.calculation.dbcs_language.unwrap_or(true), order);
    crate::recalculation::remember_completed_calculation(workbook, &report);
    if !report.deferred {
        workbook.recalculation.changed.clear();
        workbook.recalculation.dirty.clear();
        workbook.calculation.completed = None;
    }
}

/// Full-engine pass with visited coordinates for hosts retaining dirty state.
pub fn evaluate_workbook_formulas_report(workbook: &mut Workbook, now: Option<f64>) -> Vec<(String, (u32, u32))> {
    let names = workbook.defined_names.clone();
    let external = workbook.external_books.clone();
    let report = recalculate(&mut workbook.sheets, &names, &external, Overwrite::All, now, workbook.calculation.compatibility_version, workbook.calculation.dbcs_language.unwrap_or(true), &workbook.recalculation.order);
    crate::recalculation::remember_completed_calculation(workbook, &report);
    report.calculated
}

/// Recalculate selected formula cells, or formulas reached by changed cells.
/// Coordinates are zero-based. Other formula values remain cached.
pub fn evaluate_workbook_formula_subset(workbook: &mut Workbook, cells: &[(String, (u32, u32))], changed: bool, now: Option<f64>) -> Vec<(String, (u32, u32))> {
    let previous = crate::recalculation::spill_layout(workbook);
    let names = workbook.defined_names.clone(); let external = workbook.external_books.clone();
    let report = recalculate(&mut workbook.sheets, &names, &external, Overwrite::Selected { cells, changed }, now, workbook.calculation.compatibility_version, workbook.calculation.dbcs_language.unwrap_or(true), &workbook.recalculation.order);
    crate::recalculation::remember_completed_calculation(workbook, &report);
    crate::recalculation::complete(workbook, &report.calculated, &previous, false);
    report.calculated
}

/// Normal dirty calculation can reach newly filled or emptied spill cells.
/// A suspended pass returns newly discovered continuation cells; a completed
/// pass returns the learned traversal order.
pub fn evaluate_workbook_formula_dirty(workbook: &mut Workbook, cells: &[(String, (u32, u32))], now: Option<f64>, sheet_scope: Option<&str>, defer_dynamic_references: bool) -> Vec<(String, (u32, u32))> {
    let previous = crate::recalculation::spill_layout(workbook);
    let names = workbook.defined_names.clone(); let external = workbook.external_books.clone();
    let report = recalculate(&mut workbook.sheets, &names, &external, Overwrite::Dirty { cells, sheet_scope, defer_dynamic_references }, now, workbook.calculation.compatibility_version, workbook.calculation.dbcs_language.unwrap_or(true), &workbook.recalculation.order);
    crate::recalculation::remember_completed_calculation(workbook, &report);
    crate::recalculation::complete(workbook, &report.calculated, &previous, false);
    if report.deferred { report.calculated } else { report.traversal_order }
}

/// Calculate freshly entered formulas without normal-chain provisional attempts.
pub fn evaluate_workbook_formula_entry(workbook: &mut Workbook, cells: &[(String, (u32,u32))], now: Option<f64>) -> Vec<(String,(u32,u32))> {
    let names = workbook.defined_names.clone(); let external = workbook.external_books.clone();
    let report = recalculate(&mut workbook.sheets, &names, &external, Overwrite::Entry { cells }, now, workbook.calculation.compatibility_version, workbook.calculation.dbcs_language.unwrap_or(true), &workbook.recalculation.order);
    crate::recalculation::remember_completed_calculation(workbook, &report); report.calculated
}

/// Work one formula out against the workbook as it stands, without storing it
/// anywhere. `sheet` is where a bare reference like `A1` is read from.
///
/// The whole book has to be handed to the engine before it can look a
/// reference up, so this costs a pass over every cell — the same pass a
/// recalculation makes. It is for the odd expression a macro evaluates, not
/// for a loop over one.
///
/// `None` means the text is not a formula at all. A formula that fails comes
/// back as the error value a cell would show, which is what
/// `Application.Evaluate` hands a macro.
pub fn evaluate_expression(
    workbook: &Workbook,
    sheet: usize,
    formula: &str,
    now: Option<f64>,
) -> Option<oxicells_calc::Value> {
    let name = workbook.sheets.get(sheet)?.name.clone();
    let mut book = assemble_with_links(
        &workbook.sheets,
        &workbook.defined_names,
        &workbook.external_books,
        now,
    );
    // A formula cell read here answers what it holds: measured,
    book.set_compatibility_version(workbook.calculation.compatibility_version);
    book.set_dbcs_language(workbook.calculation.dbcs_language.unwrap_or(true));
    // `Evaluate("ISNA(A4)")` with A4 `=NA()` is True.
    for sheet in &workbook.sheets {
        for row in &sheet.rows {
            for cell in row.cells.iter().filter(|cell| cell.formula.is_some()) {
                let _ = book.set_cached(&sheet.name, &a1(cell.col, row.index), to_calc(&cell.value));
            }
        }
    }
    book.evaluate(&name, formula).ok()
}

/// The same, keeping a block of answers whole, as (width, height, cells row
/// by row): what `Evaluate("{1,2,3}*2")` hands a macro is three numbers.
pub fn evaluate_expression_block(
    workbook: &Workbook, sheet: usize, formula: &str, now: Option<f64>,
) -> Option<(usize, usize, Vec<oxicells_calc::Value>, bool)> {
    match evaluate_expression_argument(workbook, sheet, formula, now)? {
        oxicells_calc::functions::Arg::Value(value) => Some((1, 1, vec![value], false)),
        oxicells_calc::functions::Arg::Range(block) => {
            let array = block.is_array();
            Some((block.width, block.height, block.cells, array))
        }
    }
}

/// Evaluate once, retaining reference identity and array-value provenance.
pub fn evaluate_expression_argument(
    workbook: &Workbook,
    sheet: usize,
    formula: &str,
    now: Option<f64>,
) -> Option<oxicells_calc::functions::Arg> {
    let name = workbook.sheets.get(sheet)?.name.clone();
    let mut book = assemble_with_links(
        &workbook.sheets,
        &workbook.defined_names,
        &workbook.external_books,
        now,
    );
    book.set_compatibility_version(workbook.calculation.compatibility_version);
    book.set_dbcs_language(workbook.calculation.dbcs_language.unwrap_or(true));
    for sheet in &workbook.sheets {
        for row in &sheet.rows {
            for cell in row.cells.iter().filter(|cell| cell.formula.is_some()) {
                let _ = book.set_cached(&sheet.name, &a1(cell.col, row.index), to_calc(&cell.value));
            }
        }
    }
    book.evaluate_argument(&name, formula).ok()
}

/// The sheet (by place) and cells a formula names when it works out to a
/// reference, such as `INDEX(A1:A6,2)`; rows and columns counted from zero.
pub fn evaluate_expression_reference(
    workbook: &Workbook,
    sheet: usize,
    formula: &str,
    now: Option<f64>,
) -> Option<(usize, oxicells_calc::RangeRef)> {
    let name = workbook.sheets.get(sheet)?.name.clone();
    let book = assemble_with_links(
        &workbook.sheets,
        &workbook.defined_names,
        &workbook.external_books,
        now,
    );
    let (target, range) = book.evaluate_reference(&name, formula)?;
    let at = workbook.sheets.iter().position(|held| held.name.eq_ignore_ascii_case(&target))?;
    Some((at, range))
}

/// Compute only those formula cells the file left without a cached value,
/// leaving everything Excel already calculated untouched.
pub fn fill_missing_formula_values(workbook: &mut Workbook) {
    let names = workbook.defined_names.clone();
    let external = workbook.external_books.clone();
    recalculate(&mut workbook.sheets, &names, &external, Overwrite::OnlyMissing, None, workbook.calculation.compatibility_version, workbook.calculation.dbcs_language.unwrap_or(true), &workbook.recalculation.order);
}

/// Put the workbook to the engine: its names, its sheets, its tables and
/// every cell, either as the formula it holds or as the value it holds.
/// Put the workbook to the engine, plus what this file remembers of the
/// workbooks it links to.
///
/// A link is numbered by its place: the first `<externalReference>` is `[1]`.
/// A sheet inside it is named, and a cell of that sheet is where it was.
fn assemble_with_links(
    sheets: &[Sheet],
    names: &[(String, String)],
    external: &[crate::ir::ExternalBook],
    now: Option<f64>,
) -> oxicells_calc::Workbook {
    let mut book = assemble_sheets(sheets, names, now);
    for (at, linked) in external.iter().enumerate() {
        let number = at as u32 + 1;
        for sheet in &linked.sheets {
            for (row, column, value) in &sheet.cells {
                // The IR counts a row from one and the engine from zero.
                book.add_linked_cell(
                    number,
                    &sheet.name,
                    *column,
                    row.saturating_sub(1),
                    to_calc(value),
                );
            }
        }
    }
    book
}

fn assemble_sheets(
    sheets: &[Sheet],
    names: &[(String, String)],
    now: Option<f64>,
) -> oxicells_calc::Workbook {
    let mut book = oxicells_calc::Workbook::new();
    if let Some(moment) = now {
        book.set_now(moment);
    }

    // Named before anything is asked of them: a formula saying `SUM(sales)`
    // means one of these, and a name that will not parse is simply left
    // undefined, which is what it already was.
    for (name, refers_to) in names {
        let _ = book.define_name(name, refers_to);
    }

    for sheet in sheets.iter() {
        book.add_sheet(&sheet.name);
        // What CELL asks of a cell beyond its value: the widths of the
        // columns and each cell's number format and alignment.
        book.set_column_widths(&sheet.name, sheet.col_widths.clone(), sheet.default_col_width);
        book.set_custom_column_widths(&sheet.name, sheet.col_custom_widths.clone());
        book.set_column_width_state(&sheet.name, sheet.default_col_width_explicit, sheet.hidden_cols.iter().copied().collect());
        for row in &sheet.rows {
            for cell in &row.cells {
                if cell.style.number_format.is_some() || cell.style.horizontal_align.is_some() {
                    book.set_cell_dress(
                        &sheet.name,
                        cell.col,
                        row.index.saturating_sub(1),
                        cell.style.number_format.as_deref(),
                        cell.style.horizontal_align.as_deref(),
                    );
                }
            }
        }
        // A table's own name is how a formula reaches its columns:
        // `tblNomina[[#This Row],[DATE]]`. The IR counts a table's rows from
        // ONE and its columns from zero; the engine counts both from zero, so
        // the rows have to be shifted down on the way across. Handing them
        // over as they stand puts every table one row lower than it is, which
        // reads the headings as data and loses the last row.
        for table in &sheet.tables {
            if table.name.is_empty() {
                continue;
            }
            book.add_table(
                &sheet.name,
                &table.name,
                (
                    table.start_row.saturating_sub(1),
                    table.end_row.saturating_sub(1),
                ),
                (table.start_col, table.end_col),
                table.header_rows,
                table.totals_rows,
                table.columns.clone(),
            );
        }
        // A hidden row inside the filter's range, below its heading, is one
        // the filter hid.
        let table_filters: Vec<_> = sheet.tables.iter().filter(|table| table.has_filter_criteria())
            .map(|table| (table.start_row + table.header_rows, table.end_row.saturating_sub(table.totals_rows))).collect();
        let filtered = |index: u32| {
            sheet
                .auto_filter
                .as_ref()
                .is_some_and(|filter| index > filter.start_row && index <= filter.end_row)
                || table_filters.iter().any(|(first, last)| (*first..=*last).contains(&index))
        };
        for row in &sheet.rows {
            if row.hidden {
                book.hide_row(&sheet.name, row.index.saturating_sub(1), filtered(row.index));
            }
        }
        for row in &sheet.rows {
            for cell in &row.cells {
                let addr = a1(cell.col, row.index);
                // A member of an array formula is worked out as its share of
                // the block; a formula of its own as itself.
                let placed = match (cell.formula_units(), formula_text(cell.formula.as_deref()), cell.array_block) {
                    (Some(units), _, Some((top, left, _, _))) => book
                        .set_array_member_utf16(&sheet.name, &addr, units, (cell.col.saturating_sub(left), row.index.saturating_sub(top)))
                        .is_ok(),
                    (Some(units), _, None) => book.set_formula_utf16(&sheet.name, &addr, units).is_ok(),
                    (None, Some(text), Some((top, left, _, _))) => book
                        .set_array_member(
                            &sheet.name,
                            &addr,
                            text,
                            (cell.col.saturating_sub(left), row.index.saturating_sub(top)),
                        )
                        .is_ok(),
                    // A formula we cannot parse falls back to its cached value,
                    // so one bad cell does not poison everything downstream.
                    //
                    // Worth knowing when measuring: a formula that will not
                    // parse therefore keeps the answer the file was saved with
                    // and looks, to anything comparing the two, like a perfect
                    // agreement. Improving the parser can LOWER a measured
                    // score by letting formulas through to a gap behind them.
                    // `unparsed_formulas` counts them so that cannot hide.
                    (None, Some(text), None) => book.set_formula(&sheet.name, &addr, text).is_ok(),
                    (None, None, _) => false,
                };
                if placed && cell.spill == crate::ir::Spill::Dynamic {
                    let _ = book.set_dynamic(&sheet.name, &addr);
                }
                // A spilled value belongs to the formula that spilled it; it
                // is remembered, not placed, so it does not stand in its way.
                if !placed && cell.spill == crate::ir::Spill::Member {
                    let _ = book.set_spilled_value(&sheet.name, &addr, to_calc(&cell.value));
                    continue;
                }
                if !placed {
                    let _ = book.set_value(&sheet.name, &addr, to_calc(&cell.value));
                }
            }
        }
        // A completed calculation records each anchor's own extent. Adjacent
        // spills may touch along either axis without sharing an owner.
        if let Some(blocks) = &sheet.cached_spills {
            for ((col, row), (width, height)) in blocks {
                let Some(index) = row.checked_add(1) else { continue; };
                if *col > oxicells_calc::reference::MAX_COL || *row > oxicells_calc::reference::MAX_ROW { continue; }
                let dynamic = sheet.rows.iter().find(|held| held.index == index)
                    .and_then(|held| held.cells.iter().find(|cell| cell.col == *col))
                    .is_some_and(|cell| cell.formula.is_some() && cell.spill == crate::ir::Spill::Dynamic);
                if dynamic { let _ = book.set_cached_spill(&sheet.name, &a1(*col, index), *width, *height); }
            }
            continue;
        }
        // Older serialized workbooks do not carry the owners. Restore their
        // cached geometry using the legacy member layout until calculated.
        // Cached spill members belong to their existing dynamic anchor.
        // Restore that ownership before a subset pass; rebuilding only the
        // selected formulas must not erase unrelated spills or reduce A1#
        // to its anchor. Use the same contiguous axes as the VBA range host.
        let members: std::collections::BTreeSet<(u32, u32)> = sheet.rows.iter()
            .flat_map(|row| row.cells.iter().filter(|cell| cell.spill == crate::ir::Spill::Member)
                .map(move |cell| (cell.col, row.index))).collect();
        for row in &sheet.rows {
            for cell in row.cells.iter().filter(|cell| cell.spill == crate::ir::Spill::Dynamic) {
                let mut right = cell.col;
                let mut bottom = row.index;
                while right < oxicells_calc::reference::MAX_COL && members.contains(&(right + 1, row.index)) { right += 1; }
                while bottom <= oxicells_calc::reference::MAX_ROW && members.contains(&(cell.col, bottom + 1)) { bottom += 1; }
                if right > cell.col || bottom > row.index {
                    let _ = book.set_cached_spill(&sheet.name, &a1(cell.col, row.index), right - cell.col + 1, bottom - row.index + 1);
                }
            }
        }
    }

    book
}

#[derive(Clone, Copy, PartialEq, Eq)]
enum Overwrite<'a> {
    All,
    OnlyMissing,
    Dirty { cells: &'a [(String, (u32, u32))], sheet_scope: Option<&'a str>, defer_dynamic_references: bool },
    Selected { cells: &'a [(String, (u32, u32))], changed: bool },
    Entry { cells: &'a [(String, (u32,u32))] },
}

/// Whether a formula reads a workbook this file remembers nothing of.
///
/// Excel keeps a copy of a linked workbook so a formula reading one still has
/// an answer while the source is closed. Where it could not refresh that copy
/// it leaves an empty one, and then there is no answer to be had — so the cell
/// keeps whatever it was saved showing rather than being overwritten with
/// `#REF!`. One corpus workbook has 197 such formulas, all carrying real
/// values; answering them would throw all 197 away.
fn reads_a_forgotten_link(formula: &str, forgotten: &[u32]) -> bool {
    forgotten
        .iter()
        .any(|number| formula.contains(&format!("[{number}]")))
}

fn recalculate(
    sheets: &mut [Sheet],
    names: &[(String, String)],
    external: &[crate::ir::ExternalBook],
    mode: Overwrite,
    now: Option<f64>,
    compatibility_version: u8,
    dbcs_language: bool,
    calculation_order: &[(String, (u32, u32))],
) -> oxicells_calc::RecalcReport {
    // Which formula cells the file left without an answer. A workbook Excel
    // wrote has none, and then there is nothing to do at all — which is worth
    // asking BEFORE building anything, since building it is the cost.
    let missing: Vec<(String, (u32, u32))> = if mode == Overwrite::OnlyMissing {
        let mut found = Vec::new();
        for sheet in sheets.iter() {
            for row in &sheet.rows {
                for cell in &row.cells {
                    if formula_text(cell.formula.as_deref()).is_some()
                        && matches!(cell.value, CellValue::Empty)
                    {
                        found.push((sheet.name.clone(), (cell.col, row.index.saturating_sub(1))));
                    }
                }
            }
        }
        if found.is_empty() {
            return Default::default();
        }
        found
    } else {
        Vec::new()
    };

    let mut book = assemble_with_links(sheets, names, external, now);
    book.set_compatibility_version(compatibility_version);
    book.set_dbcs_language(dbcs_language);
    book.set_calculation_order(calculation_order);
    book.set_chain_probes(!matches!(mode, Overwrite::OnlyMissing | Overwrite::Entry { .. }));
    let forgotten: Vec<u32> = (1..=external.len() as u32)
        .filter(|number| !book.remembers_link(*number))
        .collect();

    let report = match mode {
        Overwrite::All | Overwrite::Dirty { .. } | Overwrite::Selected { .. } | Overwrite::Entry { .. } => {
            // What each formula last showed goes in with it, which a cell
            // caught in a cycle keeps: measured, `=B1+1` in A1 reads 1 and
            // still 1 once `=A1+1` in B1 closes the loop.
            for sheet in sheets.iter() {
                for row in &sheet.rows {
                    for cell in row.cells.iter().filter(|cell| cell.formula.is_some()) {
                        let _ = book.set_cached(&sheet.name, &a1(cell.col, row.index), to_calc(&cell.value));
                    }
                }
            }
            match mode {
                Overwrite::Dirty { cells, sheet_scope, defer_dynamic_references } => book.recalculate_dirty_with_dynamic_reference_wait(cells, sheet_scope, defer_dynamic_references),
                Overwrite::Selected { cells, changed: true } => book.recalculate_after(cells),
                Overwrite::Selected { cells, changed: false } => book.recalculate_these(cells),
                Overwrite::Entry { cells } => book.recalculate_these(cells),
                _ => book.recalculate(),
            }
        }
        // Only the cells that need an answer. Whatever they read is either
        // already cached or is another of these, and a missing input comes
        // earlier in the same order.
        Overwrite::OnlyMissing => {
            book.recalculate_these(&missing)
        }
    };
    let spill_blocks = book.spill_blocks();
    for sheet in sheets.iter_mut() {
        let name = sheet.name.clone();
        for row in &mut sheet.rows {
            for cell in &mut row.cells {
                if formula_text(cell.formula.as_deref()).is_none() {
                    continue;
                }
                if mode == Overwrite::OnlyMissing && !matches!(cell.value, CellValue::Empty) {
                    continue;
                }
                if !forgotten.is_empty()
                    && cell
                        .formula
                        .as_deref()
                        .is_some_and(|formula| reads_a_forgotten_link(formula, &forgotten))
                {
                    continue;
                }
                cell.value = from_calc(&book.value(&name, &a1(cell.col, row.index)));
            }
        }
        if mode != Overwrite::OnlyMissing {
            lay_out_spills(sheet, &book.spilled_values(&name));
            sheet.cached_spills = Some(spill_blocks.iter().filter(|(target, _, _)| target == &name)
                .map(|(_, anchor, extent)| (*anchor, *extent)).collect());
        }
    }
    report
}

/// Put what the formulas spilled into the cells beside them, taking back
/// whatever an earlier spill left that this one does not reach.
fn lay_out_spills(sheet: &mut Sheet, spilled: &[((u32, u32), Value)]) {
    for row in &mut sheet.rows {
        for cell in row.cells.iter_mut().filter(|cell| cell.spill == crate::ir::Spill::Member) {
            cell.value = CellValue::Empty;
            cell.spill = crate::ir::Spill::None;
        }
    }
    for ((col, row0), value) in spilled {
        let index = row0 + 1;
        let at = match sheet.rows.binary_search_by_key(&index, |row| row.index) {
            Ok(at) => at,
            Err(at) => {
                sheet.rows.insert(
                    at,
                    crate::ir::Row {
                        index,
                        cells: Vec::new(),
                        height: None,
                        custom_height: false,
                        style_font: None,
                        thick_top: false,
                        thick_bottom: false,
                        hidden: false,
                    },
                );
                at
            }
        };
        let row = &mut sheet.rows[at];
        let cell = match row.cells.iter().position(|cell| cell.col == *col) {
            Some(held) => &mut row.cells[held],
            None => {
                row.cells.push(crate::ir::Cell {
                    formula_utf16: None,
                    col: *col,
                    value: CellValue::Empty,
                    style: Default::default(),
                    formula: None,
                    runs: Vec::new(),
                    array_block: None,
                    spill: Default::default(),
                });
                row.cells.sort_by_key(|cell| cell.col);
                row.cells.iter_mut().find(|cell| cell.col == *col).expect("just put there")
            }
        };
        cell.value = from_calc(value);
        cell.spill = crate::ir::Spill::Member;
    }
}

/// A shared-formula follower carries an empty `<f>` element; treat that as
/// having no formula rather than as an empty expression.
fn formula_text(formula: Option<&str>) -> Option<&str> {
    formula.filter(|f| !f.trim().is_empty())
}

fn a1(col: u32, row_1based: u32) -> String {
    let mut out = col_to_letters(col);
    out.push_str(&row_1based.to_string());
    out
}

fn to_calc(value: &CellValue) -> Value {
    match value {
        CellValue::Empty => Value::Blank,
        CellValue::String(s) => Value::Text(s.clone()),
        CellValue::Utf16(units) => Value::Utf16(units.clone()),
        CellValue::Number(n) => Value::Number(*n),
        CellValue::Boolean(b) => Value::Logical(*b),
        CellValue::Error(s) => Value::Error(parse_error_text(s)),
    }
}

fn from_calc(value: &Value) -> CellValue {
    match value {
        Value::Blank => CellValue::Empty,
        Value::Number(n) => CellValue::Number(*n),
        Value::Text(s) => CellValue::String(s.clone()),
        Value::Utf16(units) => CellValue::from_utf16(units.clone()),
        Value::Logical(b) => CellValue::Boolean(*b),
        Value::Error(e) => CellValue::Error(e.as_str().to_string()),
    }
}

fn parse_error_text(s: &str) -> ExcelError {
    match s.to_ascii_uppercase().as_str() {
        "#NULL!" => ExcelError::Null,
        "#DIV/0!" => ExcelError::DivZero,
        "#REF!" => ExcelError::Ref,
        "#NAME?" => ExcelError::Name,
        "#NUM!" => ExcelError::Num,
        "#N/A" => ExcelError::NA,
        "#SPILL!" => ExcelError::Spill,
        "#CALC!" => ExcelError::Calc,
        _ => ExcelError::Value,
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::ir::{Cell, CellStyle, Row, Sheet};

    #[test]
    fn adjacent_dynamic_spills_keep_ownership_in_partial_calculation() {
        let mut sheet = make_sheet(vec![
            (0, 13, CellValue::Empty, Some("SEQUENCE(2,2)".into())),
            (2, 11, CellValue::Empty, Some("SEQUENCE(1,3,11)".into())),
        ]);
        for row in &mut sheet.rows { for cell in &mut row.cells { cell.spill = crate::ir::Spill::Dynamic; } }
        let name = sheet.name.clone();
        let mut workbook = Workbook { sheets: vec![sheet], ..Default::default() };
        evaluate_workbook_formulas(&mut workbook);
        for wanted in [(13, 0), (11, 2)] {
            evaluate_workbook_formula_subset(&mut workbook, &[(name.clone(), wanted)], false, None);
            assert_eq!(evaluate_expression(&workbook, 0, "ROWS(N1#)", None), Some(Value::Number(2.0)));
            assert_eq!(evaluate_expression(&workbook, 0, "COLUMNS(L3#)", None), Some(Value::Number(3.0)));
            assert_eq!(evaluate_expression(&workbook, 0, "SUM(N1#)", None), Some(Value::Number(10.0)));
            assert_eq!(evaluate_expression(&workbook, 0, "N3", None), Some(Value::Number(13.0)));
        }
    }

    #[test]
    fn explicit_nonvolatile_override_follows_only_the_original_formula() {
        let sheet = make_sheet(vec![
            (0, 0, CellValue::Number(1.0), Some("RAND()+F()".into())),
            (0, 1, CellValue::Number(1.0), Some("A1".into())),
        ]);
        let mut workbook = Workbook { sheets: vec![sheet], ..Default::default() };
        let cell = ("Sheet1".to_string(), (0, 0));
        assert!(crate::recalculation::volatile_dependents(&workbook).contains(&cell));
        workbook.recalculation.nonvolatile.insert((cell.clone(), "RAND()+F()".into()));
        assert!(crate::recalculation::volatile_dependents(&workbook).is_empty());
        workbook.sheets[0].rows[0].cells[0].formula = Some("RAND()+1".into());
        assert!(crate::recalculation::volatile_dependents(&workbook).contains(&cell));
        crate::recalculation::prepare(&mut workbook);
        assert!(workbook.recalculation.nonvolatile.is_empty());
    }

    #[test]
    fn renaming_a_sheet_keeps_pending_calculation_and_volatility_flags() {
        let sheet = make_sheet(vec![
            (0, 0, CellValue::Number(2.0), None),
            (0, 1, CellValue::Number(1.0), Some("A1+1".into())),
            (0, 2, CellValue::Number(1.0), Some("RAND()+F()".into())),
        ]);
        let mut workbook = Workbook { sheets: vec![sheet], ..Default::default() };
        workbook.recalculation.dirty.insert(("Sheet1".into(), (1, 0)));
        workbook.recalculation.nonvolatile.insert((("Sheet1".into(), (2, 0)), "RAND()+F()".into()));
        workbook.sheets[0].name = "Renamed".into();
        crate::recalculation::rename_sheet(&mut workbook, "Sheet1", "Renamed");
        crate::recalculation::prepare(&mut workbook);
        assert!(crate::recalculation::volatile_dependents(&workbook).is_empty());
        let pending: Vec<_> = workbook.recalculation.dirty.iter().cloned().collect();
        assert_eq!(pending, vec![("Renamed".to_string(), (1, 0))]);
        evaluate_workbook_formula_dirty(&mut workbook, &pending, None, None, false);
        assert_eq!(evaluate_expression(&workbook, 0, "B1", None), Some(Value::Number(3.0)));
    }

    fn make_sheet(data: Vec<(u32, u32, CellValue, Option<String>)>) -> Sheet {
        let mut rows_map: std::collections::BTreeMap<u32, Vec<Cell>> =
            std::collections::BTreeMap::new();
        for (r, c, val, formula) in data {
            rows_map.entry(r).or_default().push(Cell {
                formula_utf16: None,
                array_block: None,
                spill: Default::default(),
                col: c,
                value: val,
                style: CellStyle::default(),
                formula,
                runs: Vec::new(),
            });
        }
        let rows: Vec<Row> = rows_map
            .into_iter()
            .map(|(idx, cells)| Row {
                index: idx + 1, // 1-based
                cells,
                height: None,
                custom_height: false,
                style_font: None,
                thick_top: false,
                thick_bottom: false,
                hidden: false,
            })
            .collect();
        Sheet {
            kind: Default::default(),
            visibility: Default::default(),
            name: "Sheet1".to_string(),
            rows,
            col_count: 5,
            col_widths: vec![],
            col_custom_widths: Default::default(),
            cached_spills: None,
            pending_used_bounds: None,
            column_width_font: None,
            default_col_width: 8.43,
            default_col_width_explicit: false,
            default_row_height: 15.0,
            default_row_custom: false,
            col_fonts: vec![],
            band_styles: Default::default(),
            normal_font: None,
            first_font: None,
            frozen_rows: 0,
            frozen_cols: 0,
            merge_cells: vec![],
            hidden_cols: vec![],
            auto_filter: None,
            declared_range: None,
            tables: Vec::new(),
            drawings: Vec::new(),
            comments: Vec::new(),
            threaded_comments: Vec::new(),
            unsupported_elements: vec![],
            hyperlinks: vec![],
            protected: false,
            form_controls: Vec::new(),
            validations: Vec::new(),
            conditional_rules: Vec::new(),
            tab_color: None,
            origin_id: None,
        }
    }

    fn cross_sheet_spill_book(count: f64) -> Workbook {
        let mut source = make_sheet(vec![(0, 3, CellValue::Number(count), None), (0, 13, CellValue::Empty, Some("SEQUENCE(D1)".into()))]);
        source.name = "Source".into();source.rows[0].cells[1].spill = crate::ir::Spill::Dynamic;
        let mut target = make_sheet(vec![(0, 16, CellValue::Empty, Some("Source!N2*10".into()))]);target.name = "Target".into();
        let mut workbook = Workbook { sheets: vec![source, target], ..Default::default() };
        evaluate_workbook_formulas(&mut workbook);
        workbook.recalculation.order = vec![("Source".into(), (13, 0)), ("Target".into(), (16, 0))];
        workbook
    }

    fn change_spill_count(workbook: &mut Workbook, count: f64) {
        workbook.sheets[0].rows[0].cells[0].value = CellValue::Number(count);
        workbook.recalculation.change("Source".into(), (3, 0, 3, 0));
        crate::recalculation::prepare(workbook);
    }

    #[test]
    fn spill_growth_leaves_out_of_scope_consumers_dirty() {
        let mut workbook = cross_sheet_spill_book(1.0);
        change_spill_count(&mut workbook, 2.0);
        let target = ("Target".into(), (16, 0));
        assert!(!workbook.recalculation.dirty.contains(&target));
        let before = crate::recalculation::spill_layout(&workbook);
        let calculated = evaluate_workbook_formula_dirty(&mut workbook, &[("Source".into(), (13, 0))], None, Some("Source"), false);
        crate::recalculation::complete(&mut workbook, &calculated, &before, true);
        assert!(workbook.recalculation.dirty.contains(&target));
        assert_eq!(evaluate_expression(&workbook, 1, "Q1", None), Some(Value::Number(0.0)));
        let before = crate::recalculation::spill_layout(&workbook);
        let calculated = evaluate_workbook_formula_dirty(&mut workbook, std::slice::from_ref(&target), None, Some("Target"), false);
        crate::recalculation::complete(&mut workbook, &calculated, &before, true);
        assert_eq!(evaluate_expression(&workbook, 1, "Q1", None), Some(Value::Number(20.0)));
        assert!(workbook.recalculation.dirty.is_empty());
    }

    #[test]
    fn a_spill_wait_is_retained_across_shrinking_and_growing_again() {
        use oxicells_calc::functions::{Arg, RangeData};
        use oxicells_calc::engine::{self, UserArg};
        let mut source = make_sheet(vec![
            (0, 3, CellValue::Number(2.0), None),
            (0, 12, CellValue::Empty, Some("SIDE()".into())),
            (0, 13, CellValue::Empty, Some("ARRAY_SOURCE(D1)".into())),
        ]);
        source.name = "Source".into();
        source.rows[0].cells[2].spill = crate::ir::Spill::Dynamic;
        let mut target = make_sheet(vec![
            (0, 10, CellValue::Empty, Some("SIDE()".into())),
            (0, 16, CellValue::Empty, Some("OBSERVE(Source!N2)".into())),
        ]);
        target.name = "Target".into();
        let mut workbook = Workbook { sheets: vec![source, target], ..Default::default() };
        let observed = std::rc::Rc::new(std::cell::RefCell::new(Vec::new()));
        let seen = observed.clone();
        engine::set_user_function_hook(Some(Box::new(move |name, args, _| {
            if name == "SIDE" { return Some(Arg::Value(Value::Number(17.0))); }
            let value = match &args[0] {
                UserArg::Cells { values, .. } => values.first().cloned().unwrap_or(Value::Blank),
                UserArg::Value(value) => value.clone(),
                _ => panic!("expected scalar cell argument"),
            };
            if name == "ARRAY_SOURCE" {
                let count = value.to_number().unwrap() as usize;
                return Some(Arg::Range(RangeData {
                    origin: oxicells_calc::functions::BlockOrigin::Array,
                    width: 1, height: count,
                    cells: (1..=count).map(|i| Value::Number(i as f64)).collect(),
                }));
            }
            assert_eq!(name, "OBSERVE");
            let value = if value == Value::Blank { Value::Number(0.0) } else { value };
            seen.borrow_mut().push(value.clone());
            Some(Arg::Value(value))
        })));
        evaluate_workbook_formulas(&mut workbook);
        observed.borrow_mut().clear();
        let n = ("Source".into(), (13, 0));
        let q = ("Target".into(), (16, 0));
        let k = ("Target".into(), (10, 0));
        let m = ("Source".into(), (12, 0));
        workbook.recalculation.order = vec![q.clone(), n.clone(), k.clone(), m.clone()];
        let previous_wait = engine::user_function_wait_for_dependencies();
        engine::set_user_function_wait_for_dependencies(true);
        let mut phases = Vec::new();
        let mut chains = Vec::new();
        for count in [2.0, 1.0, 2.0] {
            change_spill_count(&mut workbook, count);
            let layout = crate::recalculation::spill_layout(&workbook);
            let previous_layout = engine::replace_user_function_spill_probe_layout(Some(layout));
            evaluate_workbook_formulas(&mut workbook);
            engine::replace_user_function_spill_probe_layout(previous_layout);
            phases.push(std::mem::take(&mut *observed.borrow_mut()));
            chains.push(workbook.recalculation.order.clone());
        }
        engine::set_user_function_wait_for_dependencies(previous_wait);
        engine::set_user_function_hook(None);
        assert_eq!(phases, vec![vec![Value::Number(0.0), Value::Number(2.0)],
            vec![Value::Number(0.0)], vec![Value::Number(2.0)]]);
        for chain in chains { assert_eq!(chain, vec![n.clone(), q.clone(), k.clone(), m.clone()]); }
    }

    #[test]
    fn temporary_traversal_order_preserves_learned_persistent_chain() {
        use oxicells_calc::functions::Arg;
        use oxicells_calc::engine::{self, UserArg};
        let mut workbook = Workbook { sheets: vec![make_sheet(vec![
            (0, 3, CellValue::Number(1.0), None),
            (0, 12, CellValue::Empty, Some("SOURCE(D1)".into())),
            (0, 13, CellValue::Empty, Some("OBSERVE(M1)".into())),
            (0, 16, CellValue::Empty, Some("SIDE(D1)".into())),
        ])], ..Default::default() };
        engine::set_user_function_hook(Some(Box::new(move |_, args, _| {
            let value = match &args[0] {
                UserArg::Cells { values, .. } => values.first().cloned().unwrap_or(Value::Blank),
                UserArg::Value(value) => value.clone(),
                _ => panic!("expected scalar cell argument"),
            };
            Some(Arg::Value(value))
        })));
        evaluate_workbook_formulas(&mut workbook);
        let m = ("Sheet1".into(), (12, 0));
        let n = ("Sheet1".into(), (13, 0));
        let q = ("Sheet1".into(), (16, 0));
        workbook.recalculation.order = vec![n.clone(), q.clone(), m.clone()];
        let temporary = vec![n.clone(), m.clone(), q.clone()];
        let previous_wait = engine::user_function_wait_for_dependencies();
        engine::set_user_function_wait_for_dependencies(true);
        evaluate_workbook_formulas_in_order(&mut workbook, &temporary, None);
        assert_eq!(workbook.recalculation.order, vec![q.clone(), m.clone(), n.clone()]);
        let original = vec![m.clone(), n.clone(), q.clone()];
        workbook.recalculation.order = original.clone();
        evaluate_workbook_formulas_in_order(&mut workbook, &[q, m, n], None);
        assert_eq!(workbook.recalculation.order, original, "a temporary traversal without repairs must not become persistent");
        engine::set_user_function_wait_for_dependencies(previous_wait);
        engine::set_user_function_hook(None);
    }

    #[test]
    fn a_scalar_wait_preserves_independent_and_downstream_formula_order() {
        use oxicells_calc::functions::Arg;
        use oxicells_calc::engine::{self, UserArg};
        // Native cases 709, 708 and 714 distinguish a clean precedent from
        // one selected alongside the consumers in a full calculation.
        for (side, tail) in [(false, false), (true, false), (false, true)] {
            let mut cells = vec![
                (0, 3, CellValue::Number(1.0), None),
                (0, 12, CellValue::Empty, Some("SOURCE(D1)".into())),
                (0, 13, CellValue::Empty, Some("OBSERVE(M1)".into())),
            ];
            if tail { cells.push((0, 14, CellValue::Empty, Some("OBSERVE(N1)".into()))); }
            if side { cells.push((0, 16, CellValue::Empty, Some("SIDE(D1)".into()))); }
            let mut workbook = Workbook { sheets: vec![make_sheet(cells)], ..Default::default() };
            let observed = std::rc::Rc::new(std::cell::RefCell::new(Vec::new()));
            let seen = observed.clone();
            engine::set_user_function_hook(Some(Box::new(move |name, args, _| {
                let value = match &args[0] {
                    UserArg::Cells { values, .. } => values.first().cloned().unwrap_or(Value::Blank),
                    UserArg::Value(value) => value.clone(),
                    _ => panic!("expected scalar cell argument"),
                }.to_number().unwrap();
                seen.borrow_mut().push((name.to_string(), value));
                Some(Arg::Value(Value::Number(value)))
            })));
            evaluate_workbook_formulas(&mut workbook);
            observed.borrow_mut().clear();
            let m = ("Sheet1".into(), (12, 0)); let n = ("Sheet1".into(), (13, 0));
            let o = ("Sheet1".into(), (14, 0)); let q = ("Sheet1".into(), (16, 0));
            let mut selected = vec![n.clone()];
            if tail { selected.push(o.clone()); } if side { selected.push(q.clone()); }
            let mut incoming = selected.clone(); incoming.push(m.clone());
            workbook.recalculation.order = incoming.clone();
            evaluate_workbook_formula_subset(&mut workbook, &selected, false, None);
            let normal = std::mem::take(&mut *observed.borrow_mut());
            let normal_order = workbook.recalculation.order.clone();
            let previous_wait = engine::user_function_wait_for_dependencies();
            engine::set_user_function_wait_for_dependencies(true);
            evaluate_workbook_formulas(&mut workbook);
            let first = std::mem::take(&mut *observed.borrow_mut());
            let learned = workbook.recalculation.order.clone();
            evaluate_workbook_formulas(&mut workbook);
            let second = std::mem::take(&mut *observed.borrow_mut());
            engine::set_user_function_wait_for_dependencies(previous_wait);
            engine::set_user_function_hook(None);
            assert_eq!(normal_order, incoming, "a clean precedent does not repair the chain");
            assert!(normal.iter().all(|(name, value)| name != "SOURCE" && *value == 1.0));
            let values: Vec<_> = first.iter().filter(|(name, _)| name == "OBSERVE").map(|(_, value)| *value).collect();
            // This hook answers synchronously; the runtime's suspended
            // callback replay is checked separately by the native corpus.
            assert_eq!(values, if tail { vec![0.0, 1.0, 1.0] } else { vec![0.0, 1.0] });
            let mut expected_order = Vec::new(); let mut expected_calls = Vec::new();
            if side { expected_order.push(q); expected_calls.push(("SIDE".to_string(), 1.0)); }
            expected_order.extend([m, n]);
            expected_calls.extend([("SOURCE".to_string(), 1.0), ("OBSERVE".to_string(), 1.0)]);
            if tail { expected_order.push(o); expected_calls.push(("OBSERVE".to_string(), 1.0)); }
            assert_eq!(learned, expected_order);
            assert_eq!(second, expected_calls, "the next full pass retains the learned order");
        }
    }

    #[test]
    fn a_scoped_consumer_with_an_unfinished_precedent_stays_dirty() {
        let mut workbook = cross_sheet_spill_book(2.0);
        change_spill_count(&mut workbook, 1.0);
        let source = ("Source".into(), (13, 0));let target = ("Target".into(), (16, 0));
        assert!(workbook.recalculation.dirty.contains(&source) && workbook.recalculation.dirty.contains(&target));
        let before = crate::recalculation::spill_layout(&workbook);
        let calculated = evaluate_workbook_formula_dirty(&mut workbook, std::slice::from_ref(&target), None, Some("Target"), false);
        crate::recalculation::complete(&mut workbook, &calculated, &before, true);
        assert_eq!(evaluate_expression(&workbook, 1, "Q1", None), Some(Value::Number(20.0)));
        assert!(workbook.recalculation.dirty.contains(&source) && workbook.recalculation.dirty.contains(&target));
        let before = crate::recalculation::spill_layout(&workbook);
        let calculated = evaluate_workbook_formula_dirty(&mut workbook, &[source, target], None, None, false);
        crate::recalculation::complete(&mut workbook, &calculated, &before, true);
        assert_eq!(evaluate_expression(&workbook, 1, "Q1", None), Some(Value::Number(0.0)));
        assert!(workbook.recalculation.dirty.is_empty());
    }

    #[test]
    fn test_simple_arithmetic() {
        let mut sheet = make_sheet(vec![
            (0, 0, CellValue::Number(10.0), None),
            (0, 1, CellValue::Number(20.0), None),
            (0, 2, CellValue::Empty, Some("A1+B1".to_string())),
        ]);
        evaluate_sheet_formulas(&mut sheet);
        assert!(
            matches!(&sheet.rows[0].cells[2].value, CellValue::Number(n) if (*n - 30.0).abs() < f64::EPSILON)
        );
    }

    #[test]
    fn test_sum_function() {
        let mut sheet = make_sheet(vec![
            (0, 0, CellValue::Number(1.0), None),
            (1, 0, CellValue::Number(2.0), None),
            (2, 0, CellValue::Number(3.0), None),
            (3, 0, CellValue::Empty, Some("SUM(A1:A3)".to_string())),
        ]);
        evaluate_sheet_formulas(&mut sheet);
        assert!(
            matches!(&sheet.rows[3].cells[0].value, CellValue::Number(n) if (*n - 6.0).abs() < f64::EPSILON)
        );
    }

    #[test]
    fn test_average_function() {
        let mut sheet = make_sheet(vec![
            (0, 0, CellValue::Number(10.0), None),
            (1, 0, CellValue::Number(20.0), None),
            (2, 0, CellValue::Number(30.0), None),
            (3, 0, CellValue::Empty, Some("AVERAGE(A1:A3)".to_string())),
        ]);
        evaluate_sheet_formulas(&mut sheet);
        assert!(
            matches!(&sheet.rows[3].cells[0].value, CellValue::Number(n) if (*n - 20.0).abs() < f64::EPSILON)
        );
    }

    #[test]
    fn test_min_max() {
        let mut sheet = make_sheet(vec![
            (0, 0, CellValue::Number(5.0), None),
            (1, 0, CellValue::Number(3.0), None),
            (2, 0, CellValue::Number(8.0), None),
            (3, 0, CellValue::Empty, Some("MIN(A1:A3)".to_string())),
            (3, 1, CellValue::Empty, Some("MAX(A1:A3)".to_string())),
        ]);
        evaluate_sheet_formulas(&mut sheet);
        assert!(
            matches!(&sheet.rows[3].cells[0].value, CellValue::Number(n) if (*n - 3.0).abs() < f64::EPSILON)
        );
        assert!(
            matches!(&sheet.rows[3].cells[1].value, CellValue::Number(n) if (*n - 8.0).abs() < f64::EPSILON)
        );
    }

    /// The old evaluator could not parse `A1>5` at all and left `#VALUE!` here.
    /// This is the single largest source of the ~23,000 wrong cells it produced
    /// across the golden-test corpus.
    #[test]
    fn test_if_with_comparison() {
        let mut sheet = make_sheet(vec![
            (0, 0, CellValue::Number(10.0), None),
            (0, 1, CellValue::Empty, Some("IF(A1>5,\"yes\",\"no\")".to_string())),
        ]);
        evaluate_sheet_formulas(&mut sheet);
        assert!(matches!(&sheet.rows[0].cells[1].value, CellValue::String(s) if s == "yes"));
    }

    #[test]
    fn test_division_by_zero() {
        let mut sheet = make_sheet(vec![
            (0, 0, CellValue::Number(10.0), None),
            (0, 1, CellValue::Number(0.0), None),
            (0, 2, CellValue::Empty, Some("A1/B1".to_string())),
        ]);
        evaluate_sheet_formulas(&mut sheet);
        assert!(matches!(&sheet.rows[0].cells[2].value, CellValue::Error(s) if s == "#DIV/0!"));
    }

    #[test]
    fn test_nested_arithmetic() {
        let mut sheet = make_sheet(vec![
            (0, 0, CellValue::Number(2.0), None),
            (0, 1, CellValue::Number(3.0), None),
            (0, 2, CellValue::Number(4.0), None),
            (0, 3, CellValue::Empty, Some("A1*B1+C1".to_string())),
        ]);
        evaluate_sheet_formulas(&mut sheet);
        assert!(
            matches!(&sheet.rows[0].cells[3].value, CellValue::Number(n) if (*n - 10.0).abs() < f64::EPSILON)
        );
    }

    #[test]
    fn test_count_function() {
        let mut sheet = make_sheet(vec![
            (0, 0, CellValue::Number(1.0), None),
            (1, 0, CellValue::String("hello".to_string()), None),
            (2, 0, CellValue::Number(3.0), None),
            (3, 0, CellValue::Empty, Some("COUNT(A1:A3)".to_string())),
        ]);
        evaluate_sheet_formulas(&mut sheet);
        assert!(
            matches!(&sheet.rows[3].cells[0].value, CellValue::Number(n) if (*n - 2.0).abs() < f64::EPSILON)
        );
    }

    /// A formula that reads another formula's result. The old snapshot-based
    /// evaluator returned the stale pre-evaluation value here.
    #[test]
    fn test_formula_chain_resolves_in_order() {
        let mut sheet = make_sheet(vec![
            (0, 0, CellValue::Number(1.0), None),
            (0, 1, CellValue::Empty, Some("A1+1".to_string())),
            (0, 2, CellValue::Empty, Some("B1+1".to_string())),
        ]);
        evaluate_sheet_formulas(&mut sheet);
        assert!(
            matches!(&sheet.rows[0].cells[2].value, CellValue::Number(n) if (*n - 3.0).abs() < f64::EPSILON)
        );
    }

    #[test]
    fn fill_missing_leaves_cached_values_alone() {
        let sheet = make_sheet(vec![
            (0, 0, CellValue::Number(1.0), None),
            // Excel cached 99 here; it disagrees with the formula on purpose.
            (0, 1, CellValue::Number(99.0), Some("A1+1".to_string())),
            (0, 2, CellValue::Empty, Some("A1+5".to_string())),
        ]);
        let mut workbook = Workbook {
            sheets: vec![sheet],
            ..Default::default()
        };
        fill_missing_formula_values(&mut workbook);
        let cells = &workbook.sheets[0].rows[0].cells;
        // Cached value kept.
        assert!(matches!(&cells[1].value, CellValue::Number(n) if (*n - 99.0).abs() < f64::EPSILON));
        // Missing value filled in.
        assert!(matches!(&cells[2].value, CellValue::Number(n) if (*n - 6.0).abs() < f64::EPSILON));
    }

    #[test]
    fn cross_sheet_references_resolve_at_workbook_level() {
        let mut data = make_sheet(vec![(0, 0, CellValue::Number(42.0), None)]);
        data.name = "Data".to_string();
        let main = make_sheet(vec![(0, 0, CellValue::Empty, Some("Data!A1*2".to_string()))]);
        let mut workbook = Workbook {
            sheets: vec![data, main],
            ..Default::default()
        };
        evaluate_workbook_formulas(&mut workbook);
        assert!(
            matches!(&workbook.sheets[1].rows[0].cells[0].value, CellValue::Number(n) if (*n - 84.0).abs() < f64::EPSILON)
        );
    }

    /// What the formula parser cannot yet read, counted over a corpus and
    /// ranked by how many WORKBOOKS each construct blocks.
    ///
    /// This is a build-order instrument, not a test: it decides what to
    /// implement next by measuring, rather than by guessing which Excel
    /// features matter. Run it with
    ///
    /// ```text
    /// cargo test -p oxicells-core formula_coverage_census -- --ignored --nocapture
    /// ```
    ///
    /// against `pipeline_data/xlsx_corpus` (see the fetch script in
    /// tools/metrics). Rank by books rather than by cells: one workbook full
    /// of the same unsupported formula is one problem, not nine hundred.
    #[test]
    #[ignore = "minutes over the whole corpus; a build-order instrument, asked for by name"]
    fn formula_coverage_census() {
        use std::collections::{BTreeMap, BTreeSet};

        let roots = [
            std::path::Path::new("../../pipeline_data/xlsx_corpus/init"),
            std::path::Path::new("../../pipeline_data/xlsx_corpus/golden"),
        ];
        let entries: Vec<_> = roots
            .iter()
            .filter_map(|root| std::fs::read_dir(root).ok())
            .flat_map(|entries| entries.flatten())
            .collect();
        if entries.is_empty() {
            println!("no corpus found; fetch it with tools/metrics/fetch_xlsx_corpus.py");
            return;
        }

        let mut books = 0usize;
        let mut blocked_books = 0usize;
        let mut formulas = 0usize;
        let mut refused = 0usize;
        let mut reasons: BTreeMap<String, (usize, usize, String)> = BTreeMap::new();

        for entry in entries {
            let path = entry.path();
            if path.extension().and_then(|e| e.to_str()) != Some("xlsx") {
                continue;
            }
            let Ok(bytes) = std::fs::read(&path) else {
                continue;
            };
            let Ok(workbook) = crate::parser::parse_xlsx_preserving_values(&bytes) else {
                continue;
            };
            books += 1;
            let mut here: BTreeSet<String> = BTreeSet::new();
            for sheet in &workbook.sheets {
                for row in &sheet.rows {
                    for cell in &row.cells {
                        let Some(formula) = cell.formula.as_deref() else {
                            continue;
                        };
                        if formula.trim().is_empty() {
                            continue;
                        }
                        formulas += 1;
                        if let Err(error) = oxicells_calc::parse(formula) {
                            refused += 1;
                            let reason = census_reason(&error.to_string(), formula);
                            here.insert(reason.clone());
                            let slot = reasons.entry(reason).or_insert((0, 0, String::new()));
                            slot.0 += 1;
                            if slot.2.is_empty() {
                                slot.2 = formula.chars().take(60).collect();
                            }
                        }
                    }
                }
            }
            if !here.is_empty() {
                blocked_books += 1;
                for reason in here {
                    reasons.entry(reason).or_default().1 += 1;
                }
            }
        }

        println!(
            "\n{books} books, {formulas} formulas, {refused} refused, {blocked_books} books blocked"
        );
        let mut ranked: Vec<_> = reasons.into_iter().collect();
        ranked.sort_by_key(|(_, (cells, books, _))| (usize::MAX - *books, usize::MAX - *cells));
        println!(
            "\n{:<34} {:>7} {:>7}  example",
            "what it could not read", "cells", "books"
        );
        for (reason, (cells, books, example)) in ranked {
            println!("{reason:<34} {cells:>7} {books:>7}  {example}");
        }
    }

    /// Group a refusal by the CONSTRUCT it stumbled on rather than by the
    /// parser's wording: the wording says where it stopped, the construct says
    /// what to build.
    fn census_reason(error: &str, formula: &str) -> String {
        // A bracketed NUMBER is another workbook -- `[1]Assistente!R6:S23` --
        // and it can sit anywhere, not only at the front. Ask that before the
        // structured-reference shape, which is a NAME followed by brackets.
        if error.contains("another workbook")
            || formula
                .match_indices('[')
                .any(|(at, _)| formula[at + 1..].starts_with(|c: char| c.is_ascii_digit()))
        {
            return "another workbook [1]Sheet!A1".to_string();
        }
        if formula.contains('{') {
            return "array constant {..}".to_string();
        }
        if formula.contains('[') && formula.contains(']') {
            return "structured reference Table[..]".to_string();
        }
        if formula.contains('@') {
            return "implicit intersection @".to_string();
        }
        if error.contains("requires a cell reference on both sides") {
            return "a `:` whose side is not a plain reference".to_string();
        }
        if error.contains("trailing input") {
            return "trailing input (often the intersection space)".to_string();
        }
        if error.contains("unexpected character") {
            let shown = error.split('\'').nth(1).unwrap_or("?");
            return format!("unexpected character {shown:?}");
        }
        if error.contains("unexpected end") {
            return "unexpected end".to_string();
        }
        format!("other: {}", error.chars().take(40).collect::<String>())
    }

    /// A formula reading a linked workbook is answered from the copy this file
    /// keeps of it.
    #[test]
    fn a_linked_formula_reads_the_copy_the_file_keeps() {
        use crate::ir::{ExternalBook, ExternalSheet};

        let mut workbook = Workbook {
            sheets: vec![make_sheet(vec![(
                0,
                0,
                CellValue::Empty,
                Some("[1]BA!$B$2*2".to_string()),
            )])],
            external_books: vec![ExternalBook {
                sheets: vec![
                    ExternalSheet {
                        name: "user".to_string(),
                        cells: Vec::new(),
                    },
                    ExternalSheet {
                        name: "BA".to_string(),
                        // One-based row, zero-based column: B2.
                        cells: vec![(2, 1, CellValue::Number(21.0))],
                    },
                ],
            }],
            ..Default::default()
        };

        evaluate_workbook_formulas(&mut workbook);
        assert!(matches!(
            workbook.sheets[0].rows[0].cells[0].value,
            CellValue::Number(n) if n == 42.0
        ));
    }

    /// A link whose copy is empty has no answer to give, so the cell keeps
    /// whatever it was saved showing rather than being overwritten.
    ///
    /// Excel leaves an empty copy behind for a link it could not refresh. One
    /// corpus workbook has 197 such formulas, every one of them carrying a
    /// real value; answering them would throw all 197 away.
    #[test]
    fn a_link_with_no_copy_leaves_the_saved_value_alone() {
        use crate::ir::{ExternalBook, ExternalSheet};

        let mut workbook = Workbook {
            sheets: vec![make_sheet(vec![(
                0,
                0,
                CellValue::Number(6.5),
                Some("[1]Assistente!R6".to_string()),
            )])],
            external_books: vec![ExternalBook {
                sheets: vec![ExternalSheet {
                    name: "Assistente".to_string(),
                    cells: Vec::new(),
                }],
            }],
            ..Default::default()
        };

        evaluate_workbook_formulas(&mut workbook);
        assert!(matches!(
            workbook.sheets[0].rows[0].cells[0].value,
            CellValue::Number(n) if n == 6.5
        ));
    }

    /// And the same over real files: every formula that reads a linked
    /// workbook works out to the value the file was saved with.
    ///
    /// 566 of them across five workbooks — including whole-column lookups
    /// (`VLOOKUP(C2,[1]Lookup!A:B,2,FALSE)`), quoted sheet names
    /// (`'[1]May 2021'!$L$2:$L$29`) and the two books whose copy is empty.
    #[test]
    fn linked_formulas_in_real_workbooks_keep_their_saved_values() {
        let books = [
            "../../pipeline_data/xlsx_corpus/golden/14240__1_14240_golden.xlsx",
            "../../pipeline_data/xlsx_corpus/golden/247-24__1_247-24_golden.xlsx",
            "../../pipeline_data/xlsx_corpus/golden/54590__1_54590_golden.xlsx",
            "../../pipeline_data/xlsx_corpus/golden/56786__1_56786_golden.xlsx",
            "../../pipeline_data/xlsx_corpus/golden/13284__golden.xlsx",
        ];
        let mut checked = 0usize;
        for name in books {
            let Ok(bytes) = std::fs::read(name) else {
                continue;
            };
            let saved = crate::parser::parse_xlsx_preserving_values(&bytes).unwrap();
            let mut worked = saved.clone();
            evaluate_workbook_formulas(&mut worked);

            for (at, sheet) in saved.sheets.iter().enumerate() {
                for row in &sheet.rows {
                    for cell in &row.cells {
                        let Some(formula) = cell.formula.as_deref() else {
                            continue;
                        };
                        if !formula.contains('[') {
                            continue;
                        }
                        let after = worked.sheets[at]
                            .rows
                            .iter()
                            .find(|r| r.index == row.index)
                            .and_then(|r| r.cells.iter().find(|c| c.col == cell.col))
                            .map(|c| format!("{:?}", c.value))
                            .unwrap_or_default();
                        assert_eq!(
                            format!("{:?}", cell.value),
                            after,
                            "{name} {formula} moved when it was worked out again"
                        );
                        checked += 1;
                    }
                }
            }
        }
        if checked > 0 {
            assert!(checked >= 500, "only {checked} linked formulas were checked");
        }
    }
}
