// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.
use super::*;
use crate::reference::Reference;

impl Workbook {
    pub fn blocked_spill_ranges(&self) -> Vec<((String, (u32, u32)), (u32, u32))> {
        self.sheets.iter().flat_map(|(name, sheet)| sheet.blocked_spills.iter().map(move |(anchor, shape)| ((name.clone(), *anchor), *shape))).collect()
    }
    /// Preferred order among independent formulas; dependencies still take precedence.
    pub fn set_calculation_order(&mut self, order: &[(String, (u32, u32))]) {
        self.calculation_order = order.iter().cloned().enumerate().map(|(i, cell)| (cell, i)).collect();
    }

    /// Rebuild the traversal from written formula-cell references. A spill
    /// member is discovered during calculation, rather than becoming a written
    /// formula prerequisite merely because it currently holds a cached value.
    pub fn rebuilt_calculation_order(&self) -> Vec<FormulaCell> {
        let keys = self.formula_keys();
        let mut at: BTreeMap<&str, BTreeMap<(u32, u32), usize>> = BTreeMap::new();
        for (i, (sheet, coord)) in keys.iter().enumerate() {
            at.entry(sheet).or_default().insert(*coord, i);
        }
        let mut prerequisites = vec![BTreeSet::new(); keys.len()];
        for (i, (sheet, coord)) in keys.iter().enumerate() {
            self.each_dependency_with_spills(sheet, coord, &at, true, false, &mut |j| {
                if j != i { prerequisites[i].insert(j); }
            });
        }
        let mut state = vec![0u8; keys.len()];
        let mut order = Vec::with_capacity(keys.len());
        // Keep the preferred chain position of a consumer by placing its
        // prerequisites immediately before it. Use an explicit stack for long
        // chains; visiting nodes also bounds cyclic input.
        for root in 0..keys.len() {
            let mut walking = vec![(root, false)];
            while let Some((i, finish)) = walking.pop() {
                if finish {
                    if state[i] != 2 { state[i] = 2; order.push(keys[i].clone()); }
                } else if state[i] == 0 {
                    state[i] = 1;
                    walking.push((i, true));
                    walking.extend(prerequisites[i].iter().rev().map(|&j| (j, false)));
                }
            }
        }
        order
    }
    /// Dependency closure without evaluating formulas or invoking UDFs.
    pub fn dependent_formula_ranges(&self, changes: &[(String, RangeRef)]) -> Vec<(String, (u32, u32))> {
        let keys = self.formula_keys();
        let mut at: BTreeMap<&str, BTreeMap<(u32, u32), usize>> = BTreeMap::new();
        for (i, (sheet, coord)) in keys.iter().enumerate() { at.entry(sheet).or_default().insert(*coord, i); }
        let mut dependents = vec![BTreeSet::new(); keys.len()];
        for (i, (sheet, coord)) in keys.iter().enumerate() {
            self.each_dependency(sheet, coord, &at, false, &mut |j| { dependents[j].insert(i); });
        }
        let mut reached = BTreeSet::new();
        let mut queue = VecDeque::new();
        for (i, (sheet, coord)) in keys.iter().enumerate() {
            let own_change = changes.iter().any(|(target, range)| same_sheet(sheet, target) && range.contains(coord.0, coord.1));
            let reads_change = self.expr_at(sheet, coord).is_some_and(|expr| {
                self.dependency_expression(&expr, sheet, Some(*coord)).invalidation_references().iter().any(|reference| {
                    reference.book.is_none() && changes.iter().any(|(target, range)| {
                        same_sheet(target, reference.sheet.as_deref().unwrap_or(sheet)) && overlaps(&reference.range, range)
                    })
                })
            });
            let obstruction_changed = self.sheets.get(sheet).and_then(|held| held.blocked_spills.get(coord)).is_some_and(|(width, height)| {
                let range = RangeRef::normalised(CellRef::new(coord.0, coord.1), CellRef::new(coord.0 + width - 1, coord.1 + height - 1));
                changes.iter().any(|(target, changed)| same_sheet(target, sheet) && overlaps(&range, changed))
            });
            if own_change || reads_change || obstruction_changed { reached.insert(i); queue.push_back(i); }
        }
        while let Some(i) = queue.pop_front() {
            for &next in &dependents[i] { if reached.insert(next) { queue.push_back(next); } }
        }
        reached.into_iter().map(|i| keys[i].clone()).collect()
    }

    /// Readers of a defined name, including aliases and their cell consumers.
    /// Lexical LET/LAMBDA bindings do not read workbook names of that spelling.
    pub fn dependent_formulas_for_name(&self, name: &str) -> Vec<FormulaCell> {
        let wanted = crate::case_fold::name_key(name);
        let roots: Vec<_> = self.formula_keys().into_iter().filter_map(|(sheet, coord)| {
            let expr = self.expr_at(&sheet, &coord)?;
            self.reads_defined_name(&expr, &sheet, &wanted, &BTreeSet::new(), &mut BTreeSet::new(), 0)
                .then(|| (sheet, RangeRef::single(CellRef::new(coord.0, coord.1))))
        }).collect();
        if roots.is_empty() { Vec::new() } else { self.dependent_formula_ranges(&roots) }
    }

    fn reads_defined_name(&self, expr: &Expr, sheet: &str, wanted: &str, locals: &BTreeSet<String>, seen: &mut BTreeSet<String>, depth: usize) -> bool {
        if depth > 100 { return false; }
        match expr {
            Expr::Name(name) if !locals.contains(&crate::case_fold::name_key(name)) => {
                let Some((key, body)) = self.name_definition(name, sheet) else { return false; };
                if key == wanted { return true; }
                if !seen.insert(key.clone()) { return false; }
                let result = self.reads_defined_name(body, sheet, wanted, &BTreeSet::new(), seen, depth + 1);
                seen.remove(key);
                result
            }
            Expr::Unary { operand, .. } => self.reads_defined_name(operand, sheet, wanted, locals, seen, depth + 1),
            Expr::Binary { lhs, rhs, .. } => self.reads_defined_name(lhs, sheet, wanted, locals, seen, depth + 1)
                || self.reads_defined_name(rhs, sheet, wanted, locals, seen, depth + 1),
            Expr::Function { name, args } => {
                if !functions::is_known_function(name) && !matches!(name.as_str(), "LET" | "LAMBDA")
                    && self.reads_defined_name(&Expr::Name(name.clone()), sheet, wanted, locals, seen, depth + 1) { return true; }
                let mut scope = locals.clone();
                for (i, arg) in args.iter().enumerate() {
                    let binding = (name == "LET" && i + 1 < args.len() && i % 2 == 0) || (name == "LAMBDA" && i + 1 < args.len());
                    if binding {
                        if name == "LAMBDA" { if let Expr::Name(label) = arg { scope.insert(crate::case_fold::name_key(label)); } }
                    } else {
                        if self.reads_defined_name(arg, sheet, wanted, &scope, seen, depth + 1) { return true; }
                        if name == "LET" && i % 2 == 1 { if let Expr::Name(label) = &args[i - 1] { scope.insert(crate::case_fold::name_key(label)); } }
                    }
                }
                false
            }
            _ => false,
        }
    }

    /// Volatile builtins, independently of dynamically registered UDF cells.
    pub fn volatile_formulas(&self) -> Vec<(String, (u32, u32))> {
        self.formula_keys().into_iter().filter(|(sheet, coord)| {
            let Some(expr) = self.expr_at(sheet, coord) else { return false };
            let resolved = self.dependency_expression(&expr, sheet, Some(*coord));
            let mut volatile = false;
            resolved.visit(&mut |node| {
                if let Expr::Function { name, args } = node {
                    if matches!(name.as_str(), "RAND" | "RANDBETWEEN" | "RANDARRAY" | "NOW" | "TODAY" | "OFFSET" | "INDIRECT" | "CELL" | "INFO" | "SHEET" | "SHEETS" | "FORMULATEXT") { volatile = true; }
                    if name == "SUMIF" {
                        if let [Expr::Ref(first), _, Expr::Ref(last)] = args.as_slice() {
                            if first.range.width() != last.range.width() || first.range.height() != last.range.height() { volatile = true; }
                        }
                    }
                }
            });
            volatile
        }).collect()
    }

    pub(super) fn dependency_expression(&self, expr: &Expr, sheet: &str, at: At) -> Expr {
        self.resolve_dependencies(expr, sheet, at, &BTreeSet::new(), &mut BTreeSet::new(), 0, false, true)
    }

    /// Written references with names and tables resolved, before volatile
    /// reference functions discover additional value prerequisites.
    pub(super) fn written_dependency_expression(&self, expr: &Expr, sheet: &str, at: At) -> Expr {
        self.resolve_dependencies(expr, sheet, at, &BTreeSet::new(), &mut BTreeSet::new(), 0, false, false)
    }

    fn resolve_dependencies(&self, expr: &Expr, sheet: &str, at: At, locals: &BTreeSet<String>, seen: &mut BTreeSet<String>, depth: usize, reference_only: bool, resolve_dynamic: bool) -> Expr {
        if depth > 100 { return expr.clone(); }
        match expr {
            Expr::Ref(reference) if reference.book.is_none() => {
                let mut reference = reference.clone();
                if let Some(target) = &reference.sheet {
                    if let Some(actual) = self.sheets.keys().find(|actual| same_sheet(actual, target)) { reference.sheet = Some(actual.clone()); }
                }
                Expr::Ref(reference)
            }
            Expr::Name(name) if !locals.contains(&crate::case_fold::name_key(name)) => {
                let key = crate::case_fold::name_key(name);
                if !seen.insert(key.clone()) { return expr.clone(); }
                let result = if let Some(bound) = self.name_bound(name, sheet, at) {
                    self.resolve_dependencies(&bound, sheet, at, locals, seen, depth + 1, reference_only, resolve_dynamic)
                } else if let Ok((target, range)) = self.table_range(name, "", sheet, at) {
                    Expr::Ref(Reference { book: None, sheet: Some(target), range })
                } else { expr.clone() };
                seen.remove(&key);
                result
            }
            Expr::Table { name, asked } => match self.table_range(name, asked, sheet, at) {
                Ok((target, range)) => Expr::Ref(Reference { book: None, sheet: Some(target), range }),
                Err(_) => expr.clone(),
            },
            Expr::Unary { op, operand } => Expr::Unary { op: *op, operand: Box::new(self.resolve_dependencies(operand, sheet, at, locals, seen, depth + 1, false, resolve_dynamic)) },
            Expr::Binary { op, lhs, rhs } => Expr::Binary { op: *op, lhs: Box::new(self.resolve_dependencies(lhs, sheet, at, locals, seen, depth + 1, false, resolve_dynamic)), rhs: Box::new(self.resolve_dependencies(rhs, sheet, at, locals, seen, depth + 1, false, resolve_dynamic)) },
            Expr::Function { name, args } => {
                let mut scope = locals.clone();
                let mut resolved = Vec::with_capacity(args.len());
                for (i, arg) in args.iter().enumerate() {
                    let binding = (name == "LET" && i + 1 < args.len() && i % 2 == 0) || (name == "LAMBDA" && i + 1 < args.len());
                    if binding {
                        if name == "LAMBDA" { if let Expr::Name(label) = arg { scope.insert(crate::case_fold::name_key(label)); } }
                        resolved.push(arg.clone());
                    } else {
                        let reference_input = user_reference_argument(name, i) || (reference_only && match name.as_str() {
                            "IF" | "CHOOSE" => i > 0,
                            "IFS" => i % 2 == 1,
                            "SWITCH" => i > 0 && (i % 2 == 0 || i + 1 == args.len()),
                            "INDEX" | "TAKE" | "DROP" | "OFFSET" | "TRIMRANGE" => i == 0,
                            "XLOOKUP" => i == 2,
                            "_UNION" | "_ISECT" | "_SPILL" => true,
                            "LET" => i + 1 == args.len(),
                            _ => false,
                        });
                        let address_only_indirect = user_function_wait_for_dependencies()
                            && user_function_address_only_argument(name, i) && {
                            let mut contains_indirect = false;
                            arg.visit(&mut |node| {
                                if matches!(node, Expr::Function { name, .. } if name == "INDIRECT") { contains_indirect = true; }
                            });
                            contains_indirect
                        };
                        let resolve_dynamic_argument = resolve_dynamic && !address_only_indirect;
                        resolved.push(self.resolve_dependencies(arg, sheet, at, &scope, seen, depth + 1, reference_input, resolve_dynamic_argument));
                        if name == "LET" && i % 2 == 1 { if let Expr::Name(label) = &args[i - 1] { scope.insert(crate::case_fold::name_key(label)); } }
                    }
                }
                let original = Expr::Function { name: name.clone(), args: resolved.clone() };
                if resolve_dynamic && !reference_only && name == "INDIRECT" && (1..=2).contains(&resolved.len())
                    && resolved.iter().all(|arg| self.dependency_value_is_fixed(arg, sheet)) {
                    if let Ok((target, range)) = self.indirect_range(&resolved, sheet, 0, at) {
                        // Preserve the volatile expression and its text inputs,
                        // while ordering the value it selects from this view.
                        return Expr::Function { name: "_UNION".into(), args: vec![original,
                            Expr::Ref(Reference { book: None, sheet: Some(target), range })] };
                    }
                }
                original
            }
            _ => expr.clone(),
        }
    }
}

fn same_sheet(a: &str, b: &str) -> bool { crate::case_fold::name_key(a) == crate::case_fold::name_key(b) }
fn overlaps(a: &RangeRef, b: &RangeRef) -> bool {
    a.start.col <= b.end.col && a.end.col >= b.start.col && a.start.row <= b.end.row && a.end.row >= b.start.row
}
