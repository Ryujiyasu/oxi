// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! Excel's Solver: the model a workbook keeps in its `solver_*` names, and
//! the answer Solver finds for it.
//!
//! Solver stores its model as sheet-local defined names: `solver_opt` (the
//! objective cell), `solver_typ` (1 maximise, 2 minimise, 3 reach
//! `solver_val`), `solver_adj` (the cells it may change), `solver_num`
//! constraints `solver_lhsN` / `solver_relN` (1 `<=`, 2 `=`, 3 `>=`, 4 int,
//! 5 bin, 6 alldifferent) / `solver_rhsN`, `solver_neg` (1: variables
//! without a bound of their own are non-negative) and `solver_eng` (1 GRG
//! Nonlinear, 2 Simplex LP, 3 Evolutionary).
//!
//! Every evaluation writes the variable cells and recalculates the whole
//! workbook, so a model is whatever the sheet computes -- the same contract
//! Solver itself has with Excel. Simplex LP reads the linear coefficients off
//! one recalculation per variable and solves the program exactly; integer and
//! binary constraints are settled by branch and bound. GRG Nonlinear is an
//! augmented-Lagrangian method over a quasi-Newton inner solve with
//! finite-difference gradients.

use crate::ir::{Cell, CellValue, Row, Workbook};

/// Which algorithm a model asks for.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Engine {
    GrgNonlinear,
    SimplexLp,
    Evolutionary,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Goal {
    Max,
    Min,
    Value(u64),
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Relation {
    Le,
    Eq,
    Ge,
    Int,
    Bin,
    AllDifferent,
}

/// A cell, zero-based (sheet index, row, column).
pub type At = (usize, u32, u32);

#[derive(Debug, Clone)]
pub enum Rhs {
    Number(f64),
    Cells(Vec<At>),
}

#[derive(Debug, Clone)]
pub struct Constraint {
    pub lhs: Vec<At>,
    pub relation: Relation,
    pub rhs: Rhs,
}

#[derive(Debug, Clone)]
pub struct Model {
    pub objective: At,
    pub goal: Goal,
    pub target: f64,
    pub variables: Vec<At>,
    pub constraints: Vec<Constraint>,
    pub non_negative: bool,
    pub engine: Engine,
}

/// Solver's `SolverSolve` return codes, as Excel numbers them.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Outcome {
    /// 0: found a solution; all constraints and optimality conditions met.
    Optimal = 0,
    /// 1: converged to the current solution.
    Converged = 1,
    /// 4: the objective cell's values do not converge (unbounded).
    Unbounded = 4,
    /// 5: could not find a feasible solution.
    Infeasible = 5,
    /// 7: the Simplex LP engine was given a model that is not linear
    /// (measured: an AllDifferent model with ABS() in its objective answers
    /// 7 under Simplex LP and leaves the cells as they were).
    NotLinear = 7,
    /// 13: the model has an error Solver will not run on.
    ModelError = 13,
}

fn sheet_of(workbook: &Workbook, name: &str) -> Option<usize> {
    let name = name.trim_matches('\'');
    workbook.sheets.iter().position(|s| s.name.eq_ignore_ascii_case(name))
}

fn col_of(letters: &str) -> Option<u32> {
    let mut col = 0u32;
    for c in letters.chars() {
        if !c.is_ascii_alphabetic() {
            return None;
        }
        col = col * 26 + (c.to_ascii_uppercase() as u32 - 'A' as u32 + 1);
    }
    col.checked_sub(1)
}

fn cell_of(text: &str) -> Option<(u32, u32)> {
    let t = text.replace('$', "");
    let letters: String = t.chars().take_while(|c| c.is_ascii_alphabetic()).collect();
    let row: u32 = t[letters.len()..].parse().ok()?;
    Some((row.checked_sub(1)?, col_of(&letters)?))
}

/// `Sheet1!$A$1:$B$2,Sheet1!$D$4` (default sheet `home`) -> cells, row-major.
fn cells_of(workbook: &Workbook, home: usize, reference: &str) -> Option<Vec<At>> {
    let reference = reference.trim().trim_start_matches('=');
    let mut out = Vec::new();
    for area in reference.split(',') {
        let (sheet, range) = match area.rsplit_once('!') {
            Some((s, r)) => (sheet_of(workbook, s)?, r),
            None => (home, area),
        };
        let (a, b) = range.split_once(':').unwrap_or((range, range));
        let (r0, c0) = cell_of(a)?;
        let (r1, c1) = cell_of(b)?;
        for r in r0.min(r1)..=r0.max(r1) {
            for c in c0.min(c1)..=c0.max(c1) {
                out.push((sheet, r, c));
            }
        }
    }
    Some(out)
}

fn number(text: &str) -> Option<f64> {
    text.trim().trim_start_matches('=').trim().parse().ok()
}

/// The Solver model a sheet keeps, if it keeps one.
pub fn read_model(workbook: &Workbook, sheet: usize) -> Option<Model> {
    let sheet_name = &workbook.sheets.get(sheet)?.name;
    let name = |leaf: &str| -> Option<String> {
        workbook.defined_names.iter().find_map(|(n, v)| {
            let (scope, rest) = n.rsplit_once('!')?;
            (rest.eq_ignore_ascii_case(leaf) && scope.trim_matches('\'').eq_ignore_ascii_case(sheet_name)).then(|| v.clone())
        })
    };
    let objective = *cells_of(workbook, sheet, &name("solver_opt")?)?.first()?;
    let goal = match name("solver_typ").and_then(|v| number(&v)).unwrap_or(1.0) as i64 {
        2 => Goal::Min,
        3 => Goal::Value(0),
        _ => Goal::Max,
    };
    let target = name("solver_val").and_then(|v| number(&v)).unwrap_or(0.0);
    let variables = cells_of(workbook, sheet, &name("solver_adj")?)?;
    let count = name("solver_num").and_then(|v| number(&v)).unwrap_or(0.0) as usize;
    let mut constraints = Vec::new();
    for k in 1..=count {
        let lhs = cells_of(workbook, sheet, &name(&format!("solver_lhs{k}"))?)?;
        let relation = match name(&format!("solver_rel{k}")).and_then(|v| number(&v)).unwrap_or(1.0) as i64 {
            2 => Relation::Eq,
            3 => Relation::Ge,
            4 => Relation::Int,
            5 => Relation::Bin,
            6 => Relation::AllDifferent,
            _ => Relation::Le,
        };
        let rhs_text = name(&format!("solver_rhs{k}")).unwrap_or_default();
        let rhs = match number(&rhs_text) {
            Some(n) => Rhs::Number(n),
            None => match cells_of(workbook, sheet, &rhs_text) {
                Some(cells) => Rhs::Cells(cells),
                None => Rhs::Number(0.0),
            },
        };
        constraints.push(Constraint { lhs, relation, rhs });
    }
    Some(Model {
        objective,
        goal,
        target,
        variables,
        constraints,
        non_negative: name("solver_neg").and_then(|v| number(&v)).unwrap_or(1.0) as i64 == 1,
        engine: match name("solver_eng").and_then(|v| number(&v)).unwrap_or(1.0) as i64 {
            2 => Engine::SimplexLp,
            3 => Engine::Evolutionary,
            _ => Engine::GrgNonlinear,
        },
    })
}

fn set_number(workbook: &mut Workbook, (s, r, c): At, value: f64) {
    let sheet = &mut workbook.sheets[s];
    let at = match sheet.rows.binary_search_by_key(&(r + 1), |row| row.index) {
        Ok(at) => at,
        Err(at) => {
            sheet.rows.insert(
                at,
                Row {
                    index: r + 1,
                    cells: Vec::new(),
                    height: None,
                    custom_height: false,
                    format_recorded: None,
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
    match row.cells.iter_mut().find(|cell| cell.col == c) {
        Some(cell) => {
            cell.value = CellValue::Number(value);
            cell.formula = None;
            cell.formula_utf16 = None;
        }
        None => {
            row.cells.push(Cell {
                col: c,
                value: CellValue::Number(value),
                style: Default::default(),
                formula: None,
                formula_utf16: None,
                runs: Vec::new(),
                array_block: None,
                spill: Default::default(),
            });
            row.cells.sort_by_key(|cell| cell.col);
        }
    }
}

fn get_number(workbook: &Workbook, (s, r, c): At) -> f64 {
    workbook.sheets[s]
        .rows
        .iter()
        .find(|row| row.index == r + 1)
        .and_then(|row| row.cells.iter().find(|cell| cell.col == c))
        .map(|cell| match &cell.value {
            CellValue::Number(n) => *n,
            CellValue::Boolean(b) => f64::from(u8::from(*b)),
            _ => 0.0,
        })
        .unwrap_or(0.0)
}

/// One pass of the model at `x`: the objective, and each constraint row as
/// (lhs, rhs) for every cell pair.
struct Probe<'a> {
    workbook: &'a mut Workbook,
    model: &'a Model,
    evaluations: usize,
}

impl Probe<'_> {
    fn eval(&mut self, x: &[f64]) -> (f64, Vec<(f64, Relation, f64)>) {
        for (cell, v) in self.model.variables.iter().zip(x) {
            set_number(self.workbook, *cell, *v);
        }
        crate::formula::evaluate_workbook_formulas(self.workbook);
        self.evaluations += 1;
        let f = get_number(self.workbook, self.model.objective);
        let mut rows = Vec::new();
        for c in &self.model.constraints {
            if matches!(c.relation, Relation::Int | Relation::Bin | Relation::AllDifferent) {
                continue;
            }
            for (i, lhs) in c.lhs.iter().enumerate() {
                let rhs = match &c.rhs {
                    Rhs::Number(n) => *n,
                    Rhs::Cells(cells) => get_number(self.workbook, cells[i.min(cells.len().saturating_sub(1))]),
                };
                rows.push((get_number(self.workbook, *lhs), c.relation, rhs));
            }
        }
        (f, rows)
    }
}

/// Which variables are integer or binary.
fn integrality(model: &Model) -> (Vec<bool>, Vec<bool>) {
    let mut int = vec![false; model.variables.len()];
    let mut bin = vec![false; model.variables.len()];
    for c in &model.constraints {
        for cell in &c.lhs {
            if let Some(i) = model.variables.iter().position(|v| v == cell) {
                match c.relation {
                    Relation::Int => int[i] = true,
                    Relation::Bin => {
                        int[i] = true;
                        bin[i] = true;
                    }
                    _ => {}
                }
            }
        }
    }
    (int, bin)
}

// ---------------------------------------------------------------- simplex --

/// min c.x subject to rows (a, rel, b), x >= lower (finite), x <= upper
/// (maybe infinite). Two-phase tableau with Bland's rule.
/// The last basic point the method stood on when it gave up: Excel leaves
/// the variable cells there (t01 infeasible and t02 unbounded both end on
/// (2, 0)), not where they started.
#[derive(Debug, Clone)]
struct Failure(Outcome, Option<Vec<f64>>);

impl From<Outcome> for Failure {
    fn from(outcome: Outcome) -> Self {
        Failure(outcome, None)
    }
}

fn simplex(c: &[f64], rows: &[(Vec<f64>, Relation, f64)], lower: &[f64], upper: &[f64]) -> Result<Vec<f64>, Failure> {
    let n = c.len();
    // Shift x = lower + y, y >= 0; add upper bounds as rows.
    let mut cons: Vec<(Vec<f64>, Relation, f64)> = Vec::new();
    for (a, rel, b) in rows {
        let shift: f64 = a.iter().zip(lower).map(|(ai, li)| ai * li).sum();
        cons.push((a.clone(), *rel, b - shift));
    }
    for j in 0..n {
        if upper[j].is_finite() {
            let mut a = vec![0.0; n];
            a[j] = 1.0;
            cons.push((a, Relation::Le, upper[j] - lower[j]));
        }
    }
    // Normalise to b >= 0.
    for (a, rel, b) in cons.iter_mut() {
        if *b < 0.0 {
            for v in a.iter_mut() {
                *v = -*v;
            }
            *b = -*b;
            *rel = match *rel {
                Relation::Le => Relation::Ge,
                Relation::Ge => Relation::Le,
                other => other,
            };
        }
    }
    let m = cons.len();
    let slacks = cons.iter().filter(|(_, r, _)| *r != Relation::Eq).count();
    let artificials = cons.iter().filter(|(_, r, _)| *r != Relation::Le).count();
    let width = n + slacks + artificials + 1;
    let mut t = vec![vec![0.0; width]; m];
    let mut basis = vec![0usize; m];
    let (mut s_at, mut a_at) = (n, n + slacks);
    let mut art_cols = Vec::new();
    for (i, (a, rel, b)) in cons.iter().enumerate() {
        t[i][..n].copy_from_slice(a);
        t[i][width - 1] = *b;
        match rel {
            Relation::Le => {
                t[i][s_at] = 1.0;
                basis[i] = s_at;
                s_at += 1;
            }
            Relation::Ge => {
                t[i][s_at] = -1.0;
                s_at += 1;
                t[i][a_at] = 1.0;
                basis[i] = a_at;
                art_cols.push(a_at);
                a_at += 1;
            }
            _ => {
                t[i][a_at] = 1.0;
                basis[i] = a_at;
                art_cols.push(a_at);
                a_at += 1;
            }
        }
    }
    let eps = 1e-9;
    let pivot = |t: &mut Vec<Vec<f64>>, basis: &mut Vec<usize>, row: usize, col: usize| {
        let p = t[row][col];
        for v in t[row].iter_mut() {
            *v /= p;
        }
        let pivot_row = t[row].clone();
        for (i, line) in t.iter_mut().enumerate() {
            if i != row && line[col].abs() > 0.0 {
                let f = line[col];
                for (v, pv) in line.iter_mut().zip(&pivot_row) {
                    *v -= f * pv;
                }
            }
        }
        basis[row] = col;
    };
    // Run the simplex on objective `obj` (minimise), columns allowed `allowed`.
    let run = |t: &mut Vec<Vec<f64>>, basis: &mut Vec<usize>, obj: &[f64], allowed: &dyn Fn(usize) -> bool| -> Result<(), Outcome> {
        for _ in 0..10_000 {
            // reduced costs
            let mut enter = None;
            for j in 0..width - 1 {
                if !allowed(j) || basis.contains(&j) {
                    continue;
                }
                let cb: f64 = (0..m).map(|i| obj[basis[i]] * t[i][j]).sum();
                if obj[j] - cb < -eps {
                    enter = Some(j);
                    break; // Bland
                }
            }
            let Some(j) = enter else { return Ok(()) };
            let mut leave = None;
            let mut best = f64::INFINITY;
            for i in 0..m {
                if t[i][j] > eps {
                    let ratio = t[i][width - 1] / t[i][j];
                    if ratio < best - eps || (ratio < best + eps && leave.is_some_and(|l: usize| basis[i] < basis[l])) {
                        best = ratio;
                        leave = Some(i);
                    }
                }
            }
            let Some(i) = leave else { return Err(Outcome::Unbounded) };
            pivot(t, basis, i, j);
        }
        Err(Outcome::Converged)
    };
    let point = |t: &Vec<Vec<f64>>, basis: &Vec<usize>| -> Vec<f64> {
        let mut y = vec![0.0; n];
        for i in 0..m {
            if basis[i] < n {
                y[basis[i]] = t[i][width - 1];
            }
        }
        y.iter().zip(lower).map(|(yi, li)| yi + li).collect()
    };
    // Phase 1.
    if !art_cols.is_empty() {
        let mut obj = vec![0.0; width];
        for a in &art_cols {
            obj[*a] = 1.0;
        }
        if let Err(outcome) = run(&mut t, &mut basis, &obj, &|_| true) {
            return Err(Failure(outcome, Some(point(&t, &basis))));
        }
        let infeasibility: f64 = (0..m).filter(|i| art_cols.contains(&basis[*i])).map(|i| t[i][width - 1]).sum();
        if infeasibility > 1e-7 {
            return Err(Failure(Outcome::Infeasible, Some(point(&t, &basis))));
        }
        // Drive artificials out of the basis where possible.
        for i in 0..m {
            if art_cols.contains(&basis[i]) {
                if let Some(j) = (0..n + slacks).find(|j| t[i][*j].abs() > eps) {
                    pivot(&mut t, &mut basis, i, j);
                }
            }
        }
    }
    let mut obj = vec![0.0; width];
    obj[..n].copy_from_slice(c);
    let arts = art_cols.clone();
    if let Err(outcome) = run(&mut t, &mut basis, &obj, &|j| !arts.contains(&j)) {
        return Err(Failure(outcome, Some(point(&t, &basis))));
    }
    Ok(point(&t, &basis))
}

fn solve_lp(probe: &mut Probe<'_>) -> Result<Vec<f64>, Failure> {
    let model = probe.model;
    let n = model.variables.len();
    let zero = vec![0.0; n];
    let (f0, rows0) = probe.eval(&zero);
    let mut c = vec![0.0; n];
    let mut a = vec![vec![0.0; n]; rows0.len()];
    for j in 0..n {
        let mut x = zero.clone();
        x[j] = 1.0;
        let (f1, rows1) = probe.eval(&x);
        c[j] = f1 - f0;
        for (i, (r1, r0)) in rows1.iter().zip(&rows0).enumerate() {
            a[i][j] = r1.0 - r0.0;
        }
    }
    // Objective as minimisation; "value of" becomes a constraint.
    let sign = match model.goal {
        Goal::Max => -1.0,
        _ => 1.0,
    };
    let cost: Vec<f64> = c.iter().map(|v| v * sign).collect();
    let mut rows: Vec<(Vec<f64>, Relation, f64)> = rows0
        .iter()
        .zip(&a)
        .map(|((l0, rel, rhs), coef)| (coef.clone(), *rel, rhs - l0))
        .collect();
    if let Goal::Value(_) = model.goal {
        rows.push((c.clone(), Relation::Eq, model.target - f0));
    }
    let cost = if matches!(model.goal, Goal::Value(_)) { vec![0.0; n] } else { cost };
    // Simplex LP asked for by name checks that the model is linear before it
    // runs: twice a unit step must move everything twice as far, and all
    // steps at once the sum of each.
    if model.engine == Engine::SimplexLp {
        let close = |a: f64, b: f64| (a - b).abs() <= 1e-6 * (1.0 + a.abs().max(b.abs()));
        for j in 0..n {
            let mut x = zero.clone();
            x[j] = 2.0;
            let (f2, rows2) = probe.eval(&x);
            if !close(f2 - f0, 2.0 * c[j]) || rows2.iter().zip(&rows0).enumerate().any(|(i, (r2, r0))| !close(r2.0 - r0.0, 2.0 * a[i][j])) {
                return Err(Failure(Outcome::NotLinear, None));
            }
        }
        let ones = vec![1.0; n];
        let (f1, rows1) = probe.eval(&ones);
        if !close(f1 - f0, c.iter().sum()) || rows1.iter().zip(&rows0).enumerate().any(|(i, (r1, r0))| !close(r1.0 - r0.0, a[i].iter().sum())) {
            return Err(Failure(Outcome::NotLinear, None));
        }
    }
    let (int, bin) = integrality(model);
    let lower = vec![if model.non_negative { 0.0 } else { -1e9 }; n];
    let upper: Vec<f64> = bin.iter().map(|b| if *b { 1.0 } else { f64::INFINITY }).collect();
    let lower: Vec<f64> = lower.iter().zip(&bin).map(|(l, b)| if *b { 0.0 } else { *l }).collect();
    if !model.non_negative {
        // Free variables: shift by a large bound is crude; split instead.
        let mut c2 = cost.clone();
        c2.extend(cost.iter().map(|v| -v));
        let rows2: Vec<_> = rows
            .iter()
            .map(|(a, r, b)| {
                let mut a2 = a.clone();
                a2.extend(a.iter().map(|v| -v));
                (a2, *r, *b)
            })
            .collect();
        let lo2 = vec![0.0; 2 * n];
        let mut up2 = upper.clone();
        up2.extend(upper.iter().map(|_| f64::INFINITY));
        let fold = |y: &[f64]| (0..n).map(|j| y[j] - y[n + j]).collect::<Vec<f64>>();
        return match branch(&c2, &rows2, &lo2, &up2, &[int.clone(), vec![false; n]].concat(), 0) {
            Ok(y) => Ok(fold(&y)),
            Err(Failure(o, p)) => Err(Failure(o, p.map(|y| fold(&y)))),
        };
    }
    branch(&cost, &rows, &lower, &upper, &int, 0)
}

/// Simplex LP's linearity check alone, for a model it will not solve as a
/// program (AllDifferent).
fn solve_lp_linearity(probe: &mut Probe<'_>) -> Result<(), Failure> {
    let n = probe.model.variables.len();
    let zero = vec![0.0; n];
    let (f0, rows0) = probe.eval(&zero);
    let close = |a: f64, b: f64| (a - b).abs() <= 1e-6 * (1.0 + a.abs().max(b.abs()));
    let mut c = vec![0.0; n];
    let mut a = vec![vec![0.0; n]; rows0.len()];
    for j in 0..n {
        let mut x = zero.clone();
        x[j] = 1.0;
        let (f1, rows1) = probe.eval(&x);
        c[j] = f1 - f0;
        for (i, (r1, r0)) in rows1.iter().zip(&rows0).enumerate() {
            a[i][j] = r1.0 - r0.0;
        }
        x[j] = 2.0;
        let (f2, rows2) = probe.eval(&x);
        if !close(f2 - f0, 2.0 * c[j]) || rows2.iter().zip(&rows0).enumerate().any(|(i, (r2, r0))| !close(r2.0 - r0.0, 2.0 * a[i][j])) {
            return Err(Failure(Outcome::NotLinear, None));
        }
    }
    let (f1, rows1) = probe.eval(&vec![1.0; n]);
    if !close(f1 - f0, c.iter().sum()) || rows1.iter().zip(&rows0).enumerate().any(|(i, (r1, r0))| !close(r1.0 - r0.0, a[i].iter().sum())) {
        return Err(Failure(Outcome::NotLinear, None));
    }
    Ok(())
}

/// AllDifferent: the N cells of each such constraint hold 1..N, each once
/// (Solver's own definition). Small groups are searched exhaustively -- up to
/// 40320 arrangements in all -- larger ones by pairwise swaps from the best
/// start until no swap helps. Variables outside the groups keep their values;
/// the other constraints must hold. Measured: an assignment of four jobs
/// whose cost has an ABS() term, under GRG Nonlinear, ends on the optimum
/// (3, 1, 4, 2) with return code 0.
fn solve_all_different(probe: &mut Probe<'_>, start: &[f64]) -> Result<Vec<f64>, Failure> {
    let model = probe.model;
    let index_of = |at: &At| model.variables.iter().position(|v| v == at);
    let mut groups: Vec<Vec<usize>> = Vec::new();
    for c in model.constraints.iter().filter(|c| c.relation == Relation::AllDifferent) {
        let Some(members) = c.lhs.iter().map(index_of).collect::<Option<Vec<usize>>>() else {
            return Err(Failure(Outcome::ModelError, None));
        };
        groups.push(members);
    }
    let sign = match model.goal {
        Goal::Max => -1.0,
        _ => 1.0,
    };
    let score = |probe: &mut Probe<'_>, x: &[f64]| -> Option<f64> {
        let (f, rows) = probe.eval(x);
        let tol = |r: f64| 1e-6 * (1.0 + r.abs());
        let feasible = rows.iter().all(|(lhs, rel, rhs)| match rel {
            Relation::Eq => (lhs - rhs).abs() <= tol(*rhs),
            Relation::Le => *lhs <= rhs + tol(*rhs),
            _ => *lhs >= rhs - tol(*rhs),
        });
        feasible.then(|| match model.goal {
            Goal::Value(_) => (f - model.target).abs(),
            _ => sign * f,
        })
    };
    let place = |perms: &[Vec<usize>]| -> Vec<f64> {
        let mut x = start.to_vec();
        for (group, perm) in groups.iter().zip(perms) {
            for (slot, value) in group.iter().zip(perm) {
                x[*slot] = (*value + 1) as f64;
            }
        }
        x
    };
    let total: f64 = groups.iter().map(|g| (1..=g.len()).map(|k| k as f64).product::<f64>()).product();
    let mut best: Option<(f64, Vec<Vec<usize>>)> = None;
    if total <= 40320.0 {
        fn permutations(n: usize) -> Vec<Vec<usize>> {
            let mut out = Vec::new();
            let mut p: Vec<usize> = (0..n).collect();
            let mut c = vec![0usize; n];
            out.push(p.clone());
            let mut i = 0;
            while i < n {
                if c[i] < i {
                    if i % 2 == 0 { p.swap(0, i) } else { p.swap(c[i], i) }
                    out.push(p.clone());
                    c[i] += 1;
                    i = 0;
                } else {
                    c[i] = 0;
                    i += 1;
                }
            }
            out
        }
        let each: Vec<Vec<Vec<usize>>> = groups.iter().map(|g| permutations(g.len())).collect();
        let mut pick = vec![0usize; groups.len()];
        loop {
            let perms: Vec<Vec<usize>> = pick.iter().zip(&each).map(|(k, all)| all[*k].clone()).collect();
            let x = place(&perms);
            if let Some(s) = score(probe, &x) {
                if best.as_ref().is_none_or(|(b, _)| s < *b - 1e-12) {
                    best = Some((s, perms));
                }
            }
            // next combination
            let mut g = 0;
            loop {
                if g == pick.len() {
                    break;
                }
                pick[g] += 1;
                if pick[g] < each[g].len() {
                    break;
                }
                pick[g] = 0;
                g += 1;
            }
            if g == pick.len() {
                break;
            }
        }
    } else {
        // Swaps from the identity until none improves.
        let mut perms: Vec<Vec<usize>> = groups.iter().map(|g| (0..g.len()).collect()).collect();
        let mut current = score(probe, &place(&perms)).unwrap_or(f64::INFINITY);
        loop {
            let mut improved = false;
            for g in 0..perms.len() {
                for i in 0..perms[g].len() {
                    for j in i + 1..perms[g].len() {
                        perms[g].swap(i, j);
                        match score(probe, &place(&perms)) {
                            Some(s) if s < current - 1e-12 => {
                                current = s;
                                improved = true;
                            }
                            _ => perms[g].swap(i, j),
                        }
                    }
                }
            }
            if !improved {
                break;
            }
        }
        if current.is_finite() {
            best = Some((current, perms));
        }
    }
    match best {
        Some((_, perms)) => Ok(place(&perms)),
        None => Err(Failure(Outcome::Infeasible, None)),
    }
}

/// Branch and bound over the integer variables.
fn branch(c: &[f64], rows: &[(Vec<f64>, Relation, f64)], lower: &[f64], upper: &[f64], int: &[bool], depth: usize) -> Result<Vec<f64>, Failure> {
    let x = simplex(c, rows, lower, upper)?;
    let frac = (0..x.len()).find(|j| int.get(*j).copied().unwrap_or(false) && (x[*j] - x[*j].round()).abs() > 1e-6);
    let Some(j) = frac else {
        return Ok(x.iter().zip(int).map(|(v, i)| if *i { v.round() } else { *v }).collect());
    };
    if depth > 60 {
        return Err(Outcome::Converged.into());
    }
    let value = |x: &[f64]| x.iter().zip(c).map(|(a, b)| a * b).sum::<f64>();
    let mut best: Option<Vec<f64>> = None;
    let mut up_low = upper.to_vec();
    up_low[j] = x[j].floor();
    if up_low[j] >= lower[j] {
        if let Ok(y) = branch(c, rows, lower, &up_low, int, depth + 1) {
            best = Some(y);
        }
    }
    let mut lo_high = lower.to_vec();
    lo_high[j] = x[j].ceil();
    if lo_high[j] <= upper[j] {
        if let Ok(y) = branch(c, rows, &lo_high, upper, int, depth + 1) {
            if best.as_ref().is_none_or(|b| value(&y) < value(b) - 1e-9) {
                best = Some(y);
            }
        }
    }
    best.ok_or(Failure(Outcome::Infeasible, None))
}

// -------------------------------------------------------------------- GRG --

fn solve_nonlinear(probe: &mut Probe<'_>, start: Vec<f64>) -> Result<Vec<f64>, Failure> {
    let model = probe.model;
    let n = model.variables.len();
    let sign = match model.goal {
        Goal::Max => -1.0,
        _ => 1.0,
    };
    let lower = if model.non_negative { 0.0 } else { f64::NEG_INFINITY };
    let project = |x: &mut Vec<f64>| {
        for v in x.iter_mut() {
            if *v < lower {
                *v = lower;
            }
        }
    };
    let mut x = start;
    project(&mut x);
    let mut lambda: Vec<f64> = Vec::new();
    let mut mu = 10.0;
    // Augmented Lagrangian of the model at x.
    let merit = |probe: &mut Probe<'_>, x: &[f64], lambda: &[f64], mu: f64| -> (f64, f64) {
        let (f, rows) = probe.eval(x);
        let mut value = match model.goal {
            Goal::Value(_) => (f - model.target).powi(2),
            _ => sign * f,
        };
        let mut worst: f64 = 0.0;
        for (k, (lhs, rel, rhs)) in rows.iter().enumerate() {
            let l = lambda.get(k).copied().unwrap_or(0.0);
            let g = lhs - rhs;
            let (term, viol) = match rel {
                Relation::Eq => (l * g + 0.5 * mu * g * g, g.abs()),
                Relation::Le => {
                    let s = (g + l / mu).max(0.0);
                    (0.5 * mu * s * s - l * l / (2.0 * mu), g.max(0.0))
                }
                _ => {
                    let s = (-g + l / mu).max(0.0);
                    (0.5 * mu * s * s - l * l / (2.0 * mu), (-g).max(0.0))
                }
            };
            value += term;
            worst = worst.max(viol);
        }
        (value, worst)
    };
    for _outer in 0..30 {
        // Inner: BFGS with projection, finite-difference gradient.
        let mut h = vec![vec![0.0; n]; n];
        for (i, row) in h.iter_mut().enumerate() {
            row[i] = 1.0;
        }
        let grad = |probe: &mut Probe<'_>, x: &[f64], lambda: &[f64], mu: f64| -> (f64, Vec<f64>) {
            // Central differences, forward where a bound is in the way.
            let (f0, _) = merit(probe, x, lambda, mu);
            let mut g = vec![0.0; n];
            for j in 0..n {
                let step = 1e-5 * x[j].abs().max(1.0);
                let mut up = x.to_vec();
                up[j] += step;
                let (f1, _) = merit(probe, &up, lambda, mu);
                if x[j] - step >= lower {
                    let mut down = x.to_vec();
                    down[j] -= step;
                    let (fm, _) = merit(probe, &down, lambda, mu);
                    g[j] = (f1 - fm) / (2.0 * step);
                } else {
                    g[j] = (f1 - f0) / step;
                }
            }
            (f0, g)
        };
        let (mut fx, mut g) = grad(probe, &x, &lambda, mu);
        let mut fresh = true;
        for _inner in 0..500 {
            let d: Vec<f64> = (0..n).map(|i| -(0..n).map(|j| h[i][j] * g[j]).sum::<f64>()).collect();
            let mut t = 1.0;
            let mut accepted = None;
            for _ in 0..40 {
                let mut y: Vec<f64> = x.iter().zip(&d).map(|(a, b)| a + t * b).collect();
                project(&mut y);
                let (fy, _) = merit(probe, &y, &lambda, mu);
                if fy < fx - 1e-12 {
                    accepted = Some((y, fy));
                    break;
                }
                t *= 0.5;
            }
            let Some((y, _)) = accepted else {
                // No descent along the quasi-Newton direction: start the
                // curvature estimate again from the gradient, once.
                if fresh {
                    break;
                }
                for (i, row) in h.iter_mut().enumerate() {
                    for (j, v) in row.iter_mut().enumerate() {
                        *v = if i == j { 1.0 / (1.0 + g.iter().map(|v| v * v).sum::<f64>().sqrt()) } else { 0.0 };
                    }
                }
                fresh = true;
                continue;
            };
            fresh = false;
            let (fy, gy) = grad(probe, &y, &lambda, mu);
            let s: Vec<f64> = y.iter().zip(&x).map(|(a, b)| a - b).collect();
            let yv: Vec<f64> = gy.iter().zip(&g).map(|(a, b)| a - b).collect();
            let sy: f64 = s.iter().zip(&yv).map(|(a, b)| a * b).sum();
            let done = (fx - fy).abs() < 1e-12 * (1.0 + fx.abs()) && s.iter().all(|v| v.abs() < 1e-9);
            if sy > 1e-12 {
                // BFGS inverse update.
                let hy: Vec<f64> = (0..n).map(|i| (0..n).map(|j| h[i][j] * yv[j]).sum()).collect();
                let yhy: f64 = yv.iter().zip(&hy).map(|(a, b)| a * b).sum();
                for i in 0..n {
                    for j in 0..n {
                        h[i][j] += (sy + yhy) * s[i] * s[j] / (sy * sy) - (hy[i] * s[j] + s[i] * hy[j]) / sy;
                    }
                }
            }
            x = y;
            fx = fy;
            g = gy;
            if done || g.iter().all(|v| v.abs() < 1e-9) {
                break;
            }
        }
        let inner_settled = g.iter().all(|v| v.abs() < 1e-6 * (1.0 + fx.abs()));
        // Multiplier update.
        let (_, rows) = probe.eval(&x);
        if lambda.len() != rows.len() {
            lambda = vec![0.0; rows.len()];
        }
        let mut worst: f64 = 0.0;
        for (k, (lhs, rel, rhs)) in rows.iter().enumerate() {
            let g = lhs - rhs;
            match rel {
                Relation::Eq => {
                    lambda[k] += mu * g;
                    worst = worst.max(g.abs());
                }
                Relation::Le => {
                    lambda[k] = (lambda[k] + mu * g).max(0.0);
                    worst = worst.max(g.max(0.0));
                }
                _ => {
                    lambda[k] = (lambda[k] - mu * g).max(0.0);
                    worst = worst.max((-g).max(0.0));
                }
            }
        }
        if worst < 1e-7 && inner_settled {
            break;
        }
        if worst >= 1e-7 {
            mu *= 4.0;
        }
    }
    let (_, rows) = probe.eval(&x);
    let feasible = rows.iter().all(|(lhs, rel, rhs)| match rel {
        Relation::Eq => (lhs - rhs).abs() < 1e-5 * (1.0 + rhs.abs()),
        Relation::Le => *lhs <= rhs + 1e-5 * (1.0 + rhs.abs()),
        _ => *lhs >= rhs - 1e-5 * (1.0 + rhs.abs()),
    });
    if feasible {
        Ok(x)
    } else {
        Err(Failure(Outcome::Infeasible, Some(x)))
    }
}

/// Solve the model on `sheet`, leave the answer in the variable cells and
/// the workbook recalculated, and say how it went.
pub fn solve(workbook: &mut Workbook, sheet: usize) -> Outcome {
    let Some(model) = read_model(workbook, sheet) else {
        return Outcome::ModelError;
    };
    let start: Vec<f64> = model.variables.iter().map(|v| get_number(workbook, *v)).collect();
    let all_different = model.constraints.iter().any(|c| c.relation == Relation::AllDifferent);
    let result = {
        let mut probe = Probe { workbook, model: &model, evaluations: 0 };
        match model.engine {
            Engine::SimplexLp if all_different => {
                // Simplex takes AllDifferent only of a linear model.
                match solve_lp_linearity(&mut probe) {
                    Ok(()) => solve_all_different(&mut probe, &start),
                    Err(failure) => Err(failure),
                }
            }
            _ if all_different => solve_all_different(&mut probe, &start),
            Engine::SimplexLp => solve_lp(&mut probe),
            _ => {
                let (int, _) = integrality(&model);
                if int.iter().any(|i| *i) {
                    // Integer models go through the linear path when they are
                    // linear; GRG with integers is branch and bound in Excel too.
                    solve_lp(&mut probe)
                } else {
                    solve_nonlinear(&mut probe, start.clone())
                }
            }
        }
    };
    match result {
        Ok(x) => {
            for (cell, v) in model.variables.iter().zip(&x) {
                let v = if v.abs() < 1e-12 { 0.0 } else { *v };
                set_number(workbook, *cell, v);
            }
            crate::formula::evaluate_workbook_formulas(workbook);
            Outcome::Optimal
        }
        Err(Failure(outcome, point)) => {
            let at = point.unwrap_or(start);
            for (cell, v) in model.variables.iter().zip(&at) {
                let v = if v.abs() < 1e-12 { 0.0 } else { *v };
                set_number(workbook, *cell, v);
            }
            crate::formula::evaluate_workbook_formulas(workbook);
            outcome
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn simplex_finds_the_textbook_vertex() {
        // max 3x + 2y, x + y <= 4, x + 3y <= 6, x, y >= 0  ->  (4, 0)
        let rows = vec![(vec![1.0, 1.0], Relation::Le, 4.0), (vec![1.0, 3.0], Relation::Le, 6.0)];
        let x = simplex(&[-3.0, -2.0], &rows, &[0.0, 0.0], &[f64::INFINITY, f64::INFINITY]).unwrap();
        assert!((x[0] - 4.0).abs() < 1e-9 && x[1].abs() < 1e-9, "{x:?}");
    }

    #[test]
    fn simplex_handles_equalities_and_lower_bounds() {
        // min x + y, x + 2y = 6, x >= 1, y >= 0 -> (1, 2.5)? min x+y with x+2y=6: y as large -> x=1? x>=1 -> y=2.5, obj 3.5
        let rows = vec![(vec![1.0, 2.0], Relation::Eq, 6.0)];
        let x = simplex(&[1.0, 1.0], &rows, &[1.0, 0.0], &[f64::INFINITY, f64::INFINITY]).unwrap();
        assert!((x[0] - 1.0).abs() < 1e-9 && (x[1] - 2.5).abs() < 1e-9, "{x:?}");
    }

    #[test]
    fn branch_and_bound_rounds_to_the_integer_optimum() {
        // max 5x + 4y, 6x + 4y <= 24, x + 2y <= 6, integers -> (4, 0) obj 20? (3,1)=19,(4,0)=20
        let rows = vec![(vec![6.0, 4.0], Relation::Le, 24.0), (vec![1.0, 2.0], Relation::Le, 6.0)];
        let x = branch(&[-5.0, -4.0], &rows, &[0.0, 0.0], &[f64::INFINITY, f64::INFINITY], &[true, true], 0).unwrap();
        assert_eq!(x, vec![4.0, 0.0]);
    }
}
