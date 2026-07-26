// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! Measure the engine against real workbooks, without needing Excel installed.
//!
//! # Why this works
//!
//! An `.xlsx` file stores, for every formula cell, both the formula *and* the
//! value Excel last computed for it. That cached value is a free oracle: load
//! the file, recalculate from the formulas alone, and diff against what Excel
//! put there. No COM, no Excel licence, no Windows.
//!
//! The output that matters most is the ranked list of unimplemented functions.
//! It says which functions to write next based on how often they actually occur
//! in real files, rather than on a guess about what is common.
//!
//! ```text
//! cargo run -p oxicells-calc --example xlsx_oracle -- <dir-or-file>...
//! ```
//! With no arguments it looks in `tools/golden-test/documents/xlsx`.

use std::collections::BTreeMap;
use std::path::{Path, PathBuf};

use oxicells_calc::ast::Expr;
use oxicells_calc::value::ExcelError;
use oxicells_calc::{parse, Value, Workbook};
use oxicells_core::ir::{self, CellValue};

/// Excel carries 15 significant digits, so anything tighter than this is noise
/// from the last binary place rather than a real disagreement.
const RELATIVE_TOLERANCE: f64 = 1e-10;

#[derive(Default)]
struct Report {
    files_read: usize,
    files_failed: usize,
    formulas: usize,
    /// Formula cells the file carries no cached value for. There is nothing to
    /// compare against, so they are excluded from the denominator rather than
    /// counted as agreement or disagreement either way.
    no_oracle: usize,
    matched: usize,
    /// Our value is an error, Excel had a real value.
    we_errored: BTreeMap<String, usize>,
    /// Both sides produced a value, but they differ.
    mismatched: usize,
    /// Shape of each disagreement, as `ours -> excel`. The single most useful
    /// number when the match rate is unexpectedly low: it says whether the
    /// engine is computing wrong values or not computing at all.
    mismatch_kinds: BTreeMap<String, usize>,
    /// Functions that appear in real formulas but are not implemented.
    missing_functions: BTreeMap<String, usize>,
    /// Formulas that would not even parse.
    unparsed: usize,
    unparsed_samples: Vec<String>,
    /// Raw `CellValue::Error` payloads seen on formula cells, verbatim. Counting
    /// these is the only way to tell a real Excel error apart from a value the
    /// reader failed to classify.
    excel_error_strings: BTreeMap<String, usize>,
    samples: Vec<String>,
}

fn type_name(v: &Value) -> &'static str {
    match v {
        Value::Blank => "blank",
        Value::Number(_) => "number",
        Value::Text(_) => "text",
        Value::Logical(_) => "logical",
        Value::Error(_) => "error",
    }
}

fn main() {
    let args: Vec<String> = std::env::args().skip(1).collect();
    let roots: Vec<PathBuf> = if args.is_empty() {
        vec![default_corpus()]
    } else {
        args.iter().map(PathBuf::from).collect()
    };

    let mut files = Vec::new();
    for root in &roots {
        collect_xlsx(root, &mut files);
    }
    files.sort();

    if files.is_empty() {
        eprintln!("no .xlsx files found under {roots:?}");
        std::process::exit(1);
    }

    let mut report = Report::default();
    for path in &files {
        match std::fs::read(path) {
            // NOT parse_xlsx: that one recalculates with oxicells-core's own
            // evaluator and overwrites the values Excel cached, which would make
            // this harness compare the engine against the thing it replaces.
            Ok(bytes) => match oxicells_core::parse_xlsx_preserving_values(&bytes) {
                Ok(book) => {
                    report.files_read += 1;
                    measure(&book, path, &mut report);
                }
                Err(_) => report.files_failed += 1,
            },
            Err(_) => report.files_failed += 1,
        }
    }

    print_report(&report, files.len());
}

fn default_corpus() -> PathBuf {
    // examples/ -> oxicells-calc/ -> crates/ -> repo root
    Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../tools/golden-test/documents/xlsx")
}

fn collect_xlsx(path: &Path, out: &mut Vec<PathBuf>) {
    if path.is_file() {
        if path.extension().is_some_and(|e| e.eq_ignore_ascii_case("xlsx")) {
            out.push(path.to_path_buf());
        }
        return;
    }
    let Ok(entries) = std::fs::read_dir(path) else {
        return;
    };
    for entry in entries.flatten() {
        collect_xlsx(&entry.path(), out);
    }
}

/// Translate the renderer's cell model into a calc value.
fn to_value(v: &CellValue) -> Value {
    match v {
        CellValue::Empty => Value::Blank,
        CellValue::String(s) => Value::Text(s.clone()),
        CellValue::Number(n) => Value::Number(*n),
        CellValue::Boolean(b) => Value::Logical(*b),
        CellValue::Error(s) => Value::Error(parse_error_text(s)),
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
        _ => ExcelError::Value,
    }
}

fn a1(col: u32, row_1based: u32) -> String {
    format!(
        "{}{}",
        oxicells_calc::reference::col_to_letters(col),
        row_1based
    )
}

fn measure(book: &ir::Workbook, path: &Path, report: &mut Report) {
    let mut wb = Workbook::new();
    // (sheet, a1) -> the value Excel cached for that formula.
    let mut expected: Vec<(String, String, Value)> = Vec::new();

    for sheet in &book.sheets {
        wb.add_sheet(&sheet.name);
        for row in &sheet.rows {
            for cell in &row.cells {
                let addr = a1(cell.col, row.index);
                match &cell.formula {
                    Some(formula) => {
                        // Collect function names before anything can fail, so
                        // that an unparsed formula still reports nothing wrong
                        // about functions.
                        if let Ok(expr) = parse(formula) {
                            record_functions(&expr, report);
                        }
                        if wb.set_formula(&sheet.name, &addr, formula).is_err() {
                            report.unparsed += 1;
                            if report.unparsed_samples.len() < 10 {
                                report.unparsed_samples.push(formula.clone());
                            }
                            continue;
                        }
                        if let CellValue::Error(raw) = &cell.value {
                            *report
                                .excel_error_strings
                                .entry(raw.clone())
                                .or_default() += 1;
                        }
                        expected.push((sheet.name.clone(), addr, to_value(&cell.value)));
                    }
                    None => {
                        let _ = wb.set_value(&sheet.name, &addr, to_value(&cell.value));
                    }
                }
            }
        }
    }

    wb.recalculate();

    let file = path.file_name().unwrap_or_default().to_string_lossy();
    for (sheet, addr, excel) in expected {
        report.formulas += 1;
        if excel.is_blank() {
            report.no_oracle += 1;
            continue;
        }
        let ours = wb.value(&sheet, &addr);
        if agrees(&ours, &excel) {
            report.matched += 1;
            continue;
        }
        *report
            .mismatch_kinds
            .entry(format!("{} -> {}", type_name(&ours), type_name(&excel)))
            .or_default() += 1;
        match ours.err() {
            Some(e) if !excel.is_error() => {
                *report.we_errored.entry(e.as_str().to_string()).or_default() += 1;
            }
            _ => report.mismatched += 1,
        }
        // At most two samples per file, so a single pathological workbook
        // cannot fill the list and hide everything else.
        let seen = report
            .samples
            .iter()
            .filter(|s| s.starts_with(file.as_ref()))
            .count();
        if report.samples.len() < 60 && seen < 2 {
            let formula = wb.formula(&sheet, &addr).unwrap_or("?");
            report.samples.push(format!(
                "{file} {sheet}!{addr}  {formula}\n      excel={excel:?}\n      ours ={ours:?}"
            ));
        }
    }
}

fn record_functions(expr: &Expr, report: &mut Report) {
    expr.visit(&mut |node| {
        if let Expr::Function { name, .. } = node {
            // An implemented function with no arguments reports #VALUE! (or a
            // value); only the registry fallthrough reports #NAME?. That makes
            // this a reliable "is it implemented" probe.
            if oxicells_calc::functions::call(name, &[]).err() == Some(ExcelError::Name) {
                *report.missing_functions.entry(name.clone()).or_default() += 1;
            }
        }
    });
}

/// Excel's cached value is rounded to 15 significant digits on the way into the
/// file, so exact float equality would report false divergences.
fn agrees(ours: &Value, excel: &Value) -> bool {
    match (ours, excel) {
        (Value::Number(a), Value::Number(b)) => {
            let scale = a.abs().max(b.abs()).max(1.0);
            (a - b).abs() <= RELATIVE_TOLERANCE * scale
        }
        // Excel writes an empty formula result as an empty string.
        (Value::Blank, Value::Text(s)) | (Value::Text(s), Value::Blank) => s.is_empty(),
        _ => ours == excel,
    }
}

fn print_report(r: &Report, total_files: usize) {
    let comparable = r.formulas - r.no_oracle;
    let pct = |n: usize| {
        if comparable == 0 {
            0.0
        } else {
            n as f64 * 100.0 / comparable as f64
        }
    };

    println!("\n=== oxicells-calc vs Excel's cached values ===\n");
    println!("files            {total_files} found / {} read / {} unreadable", r.files_read, r.files_failed);
    println!("formula cells    {}", r.formulas);
    println!("  no cached value {:>7}  (excluded: nothing to compare against)", r.no_oracle);
    println!("comparable       {comparable}");
    println!("  matched        {:>7}  ({:.3}%)", r.matched, pct(r.matched));
    println!("  diverged       {:>7}  ({:.3}%)", r.mismatched, pct(r.mismatched));
    let errored: usize = r.we_errored.values().sum();
    println!("  we errored     {:>7}  ({:.3}%)", errored, pct(errored));
    println!("  unparsed       {:>7}", r.unparsed);

    if !r.mismatch_kinds.is_empty() {
        println!("\n-- disagreement shapes (ours -> excel) --");
        let mut rows: Vec<_> = r.mismatch_kinds.iter().collect();
        rows.sort_by(|a, b| b.1.cmp(a.1));
        for (kind, count) in rows.iter().take(12) {
            println!("  {count:>7}  {kind}");
        }
    }

    if !r.excel_error_strings.is_empty() {
        println!("\n-- raw cached-error payloads on formula cells --");
        let mut rows: Vec<_> = r.excel_error_strings.iter().collect();
        rows.sort_by(|a, b| b.1.cmp(a.1));
        for (raw, count) in rows.iter().take(12) {
            println!("  {count:>7}  {raw:?}");
        }
    }

    if !r.unparsed_samples.is_empty() {
        println!("\n-- formulas that would not parse --");
        for s in &r.unparsed_samples {
            println!("  {s}");
        }
    }

    if !r.we_errored.is_empty() {
        println!("\n-- our error values where Excel had a result --");
        let mut rows: Vec<_> = r.we_errored.iter().collect();
        rows.sort_by(|a, b| b.1.cmp(a.1));
        for (kind, count) in rows {
            println!("  {count:>7}  {kind}");
        }
    }

    if !r.missing_functions.is_empty() {
        println!("\n-- unimplemented functions, by real-world frequency --");
        let mut rows: Vec<_> = r.missing_functions.iter().collect();
        rows.sort_by(|a, b| b.1.cmp(a.1).then(a.0.cmp(b.0)));
        for (name, count) in rows.iter().take(40) {
            println!("  {count:>7}  {name}");
        }
        if rows.len() > 40 {
            println!("  ... and {} more", rows.len() - 40);
        }
    }

    if !r.samples.is_empty() {
        println!("\n-- sample divergences --");
        for s in &r.samples {
            println!("  {s}");
        }
    }
    println!();
}
