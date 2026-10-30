// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.
//! Pending formula work, independent of the macro or editor that changed it.
use crate::ir::{CellValue, FormulaLocation, Workbook};
use oxicells_calc::reference::{CellRef, RangeRef};
use std::collections::{BTreeMap, BTreeSet};

/// Zero-based anchors and extents of currently published array blocks.
pub type SpillLayout = Vec<(FormulaLocation, (u32, u32))>;

pub fn spill_layout(workbook: &Workbook) -> SpillLayout {
    crate::formula::assemble_for_evaluation(&workbook.sheets, &workbook.defined_names, &workbook.external_books)
        .spill_blocks().into_iter().map(|(sheet, anchor, size)| ((sheet, anchor), size)).collect()
}

/// Finish a calculation without forgetting work outside its requested scope.
/// Newly gained/lost spill cells invalidate their consumers on every sheet.
/// A consumer calculated against an unfinished precedent remains dirty.
pub fn complete(workbook: &mut Workbook, calculated: &[FormulaLocation], previous: &SpillLayout, clear_dirty: bool) {
    let book = crate::formula::assemble_for_evaluation(&workbook.sheets, &workbook.defined_names, &workbook.external_books);
    let before: BTreeMap<_, _> = previous.iter().cloned().collect();
    let after: BTreeMap<_, _> = book.spill_blocks().into_iter().map(|(sheet, anchor, size)| ((sheet, anchor), size)).collect();
    let anchors: BTreeSet<_> = before.keys().chain(after.keys()).cloned().collect();
    let mut changes = Vec::new();
    for (sheet, anchor) in anchors {
        let old = before.get(&(sheet.clone(), anchor)).copied().unwrap_or((1, 1));
        let new = after.get(&(sheet.clone(), anchor)).copied().unwrap_or((1, 1));
        if old == new { continue; }
        // The anchor also carries the extent read through the spill operator.
        changes.push((sheet.clone(), RangeRef::single(CellRef::new(anchor.0, anchor.1))));
        for (outer, inner) in [(old, new), (new, old)] {
            if outer.0 > inner.0 {
                changes.push((sheet.clone(), RangeRef::normalised(CellRef::new(anchor.0 + inner.0, anchor.1), CellRef::new(anchor.0 + outer.0 - 1, anchor.1 + outer.1 - 1))));
            }
            if outer.1 > inner.1 {
                changes.push((sheet.clone(), RangeRef::normalised(CellRef::new(anchor.0, anchor.1 + inner.1), CellRef::new(anchor.0 + outer.0.min(inner.0) - 1, anchor.1 + outer.1 - 1))));
            }
        }
    }
    if !changes.is_empty() { workbook.recalculation.dirty.extend(book.dependent_formula_ranges(&changes)); }
    if !clear_dirty { return; }
    let done: BTreeSet<_> = calculated.iter().cloned().collect();
    let remaining: Vec<_> = workbook.recalculation.dirty.difference(&done)
        .map(|(sheet, cell)| (sheet.clone(), RangeRef::single(CellRef::new(cell.0, cell.1)))).collect();
    let unfinished: BTreeSet<_> = if remaining.is_empty() { BTreeSet::new() }
        else { book.dependent_formula_ranges(&remaining).into_iter().collect() };
    for cell in calculated {
        if !unfinished.contains(cell) { workbook.recalculation.dirty.remove(cell); }
    }
}

pub fn formulas(workbook: &Workbook) -> BTreeSet<FormulaLocation> {
    workbook.sheets.iter().flat_map(|sheet| sheet.rows.iter().flat_map(move |row| {
        row.cells.iter().filter(|cell| cell.formula.is_some()).map(move |cell| (sheet.name.clone(), (cell.col, row.index.saturating_sub(1))))
    })).collect()
}

pub fn initialize(workbook: &mut Workbook) {
    if workbook.calculation.completed == Some(false) && !workbook.recalculation.pending() {
        workbook.recalculation.dirty.extend(formulas(workbook));
    }
    for sheet in &workbook.sheets {
        for row in &sheet.rows {
            for cell in &row.cells {
                if cell.formula.is_some() && matches!(cell.value, CellValue::Empty) {
                    workbook.recalculation.dirty.insert((sheet.name.clone(), (cell.col, row.index.saturating_sub(1))));
                }
            }
        }
    }
}

/// Carry pending work and explicit UDF flags with a renamed worksheet.
pub fn rename_sheet(workbook: &mut Workbook, old: &str, new: &str) {
    let rename = |cell: &mut FormulaLocation| {
        if cell.0 == old { cell.0 = new.to_string(); }
    };
    let state = &mut workbook.recalculation;
    for cell in &mut state.order { rename(cell); }
    if let Some(order) = &mut state.forced_order { for cell in order { rename(cell); } }
    for (cell, _) in &mut state.blocked_spills { rename(cell); }
    for (sheet, _) in &mut state.changed { if sheet == old { *sheet = new.to_string(); } }
    state.dirty = std::mem::take(&mut state.dirty).into_iter().map(|mut cell| { rename(&mut cell); cell }).collect();
    state.volatile = std::mem::take(&mut state.volatile).into_iter().map(|mut cell| { rename(&mut cell); cell }).collect();
    state.nonvolatile = std::mem::take(&mut state.nonvolatile).into_iter().map(|(mut cell, formula)| {
        rename(&mut cell);
        (cell, oxicells_calc::rename_sheet_in_formula(&formula, old, new))
    }).collect();
}

/// Apply pending edit rectangles and discard marks on formulas that were removed.
pub fn prepare(workbook: &mut Workbook) {
    let present = formulas(workbook);
    workbook.recalculation.order.retain(|cell| present.contains(cell));
    if let Some(order) = &mut workbook.recalculation.forced_order { order.retain(|cell| present.contains(cell)); }
    workbook.recalculation.blocked_spills.retain(|(cell, _)| present.contains(cell));
    workbook.recalculation.dirty.retain(|cell| present.contains(cell));
    workbook.recalculation.volatile.retain(|cell| present.contains(cell));
    let sheets = &workbook.sheets;
    workbook.recalculation.nonvolatile.retain(|(cell, formula)| formula_at(sheets, cell) == Some(formula.as_str()));
    if !workbook.recalculation.changed.is_empty() {
        let book = crate::formula::assemble_for_evaluation(&workbook.sheets, &workbook.defined_names, &workbook.external_books);
        let mut changes = workbook.recalculation.changed.iter().map(|(sheet, (left, top, right, bottom))| {
            (sheet.clone(), RangeRef::normalised(CellRef::new(*left, *top), CellRef::new(*right, *bottom)))
        }).collect::<Vec<_>>();
        let unblocked: Vec<_> = workbook.recalculation.blocked_spills.iter().filter(|((sheet, (col, row)), (width, height))| {
            changes.iter().any(|(target, range)| target == sheet && range.start.col < col + width && range.end.col >= *col && range.start.row < row + height && range.end.row >= *row)
        }).map(|((sheet, (col, row)), _)| (sheet.clone(), RangeRef::single(CellRef::new(*col, *row)))).collect();
        changes.extend(unblocked);
        workbook.recalculation.dirty.extend(book.dependent_formula_ranges(&changes));
        workbook.recalculation.changed.clear();
    }
}

pub(crate) fn remember_completed_calculation(workbook: &mut Workbook, report: &oxicells_calc::RecalcReport) {
    if report.deferred { return; }
    // Defer consumers using only dependencies learned in this pass. Stable
    // ordering preserves independent formulas and siblings sharing a producer;
    // splicing each consumer separately would reverse those siblings.
    // Completion order also contains temporary computed-reference waits and
    // must not replace the persistent chain.
    if !report.dependency_order_repairs.is_empty() {
        let order = &workbook.recalculation.order;
        let positions: BTreeMap<_, _> = order.iter().enumerate().map(|(i, cell)| (cell, i)).collect();
        let mut dependents = vec![BTreeSet::new(); order.len()];
        let mut indegree = vec![0usize; order.len()];
        for (producer, consumer) in &report.dependency_order_repairs {
            if let (Some(&from), Some(&to)) = (positions.get(producer), positions.get(consumer)) {
                if from != to && dependents[from].insert(to) { indegree[to] += 1; }
            }
        }
        let mut ready: BTreeSet<_> = (0..order.len()).filter(|i| indegree[*i] == 0).collect();
        let mut repaired = Vec::with_capacity(order.len());
        while let Some(i) = ready.pop_first() {
            repaired.push(order[i].clone());
            for &consumer in &dependents[i] {
                indegree[consumer] -= 1;
                if indegree[consumer] == 0 { ready.insert(consumer); }
            }
        }
        if repaired.len() == order.len() { workbook.recalculation.order = repaired; }
    }
    let calculated: BTreeSet<_> = report.calculated.iter().cloned().collect();
    workbook.recalculation.blocked_spills.retain(|(cell, _)| !calculated.contains(cell));
    workbook.recalculation.blocked_spills.extend(report.blocked_spills.iter().cloned());
}

fn formula_at<'a>(sheets: &'a [crate::ir::Sheet], (sheet, (col, row)): &FormulaLocation) -> Option<&'a str> {
    sheets.iter().find(|item| &item.name == sheet)?
        .rows.iter().find(|item| item.index == row.saturating_add(1))?
        .cells.iter().find(|item| item.col == *col)?.formula.as_deref()
}

pub fn volatile_dependents(workbook: &Workbook) -> BTreeSet<FormulaLocation> {
    let book = crate::formula::assemble_for_evaluation(&workbook.sheets, &workbook.defined_names, &workbook.external_books);
    let suppressed: BTreeSet<_> = workbook.recalculation.nonvolatile.iter()
        .filter(|(cell, formula)| formula_at(&workbook.sheets, cell) == Some(formula.as_str()))
        .map(|(cell, _)| cell.clone()).collect();
    let roots: BTreeSet<_> = book.volatile_formulas().into_iter().filter(|cell| !suppressed.contains(cell))
        .chain(workbook.recalculation.volatile.iter().cloned()).collect();
    let changes = roots.iter().map(|(sheet, (col, row))| (sheet.clone(), RangeRef::single(CellRef::new(*col, *row)))).collect::<Vec<_>>();
    book.dependent_formula_ranges(&changes).into_iter().collect()
}
