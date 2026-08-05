// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

use std::collections::{BTreeMap, BTreeSet};
use std::fs;
use std::path::{Path, PathBuf};

use oxidocs_common::archive::OoxmlArchive;
use oxivba_core::fingerprint::{
    compare, fingerprint_module, ModuleFingerprint, Similarity, Strength,
};
use oxivba_core::{analyse, parse_module, Analysis, Class};
use serde_json::json;

struct ModuleReport {
    name: String,
    analysis: Analysis,
    fingerprint: ModuleFingerprint,
}

struct ProjectReport {
    path: PathBuf,
    container_part: String,
    modules: Vec<ModuleReport>,
}

struct RelatedModules {
    left: String,
    right: String,
    similarity: Similarity,
}

pub(crate) fn analyze_file(input: &str) -> Result<(), String> {
    let report = inspect_project(Path::new(input))?;

    println!("VBA project: {}", report.path.display());
    println!("Container part: {}", report.container_part);
    println!("Modules: {}", report.modules.len());
    println!();

    for module in &report.modules {
        println!("[{}]", module.name);
        println!("  verdict: {}", module.analysis.verdict());
        println!(
            "  procedures: {}, statements: {}, max nesting: {}, unparsed: {}",
            module.analysis.metrics.procedures,
            module.analysis.metrics.statements,
            module.analysis.metrics.max_nesting,
            module.analysis.metrics.unparsed
        );
        println!(
            "  formula engine: {}",
            if module.analysis.needs_formula_engine {
                "required"
            } else {
                "not detected"
            }
        );
        if !module.analysis.external_declares.is_empty() {
            println!(
                "  external declares: {}",
                module.analysis.external_declares.join(", ")
            );
        }
        for finding in &module.analysis.findings {
            let class = finding.class.map_or("-", |class| class.as_str());
            println!(
                "  finding line {} [{}] {}: {}",
                finding.line, class, finding.what, finding.reason
            );
        }
        println!();
    }

    let (procedures, statements, unparsed) = project_totals(&report);
    println!(
        "Summary: {} module(s), {procedures} procedure(s), {statements} statement(s), {unparsed} unparsed line(s)",
        report.modules.len()
    );
    Ok(())
}

pub(crate) fn inventory(input: &str) -> Result<(), String> {
    let root = Path::new(input);
    let mut files = Vec::new();
    collect_macro_files(root, &mut files)?;
    files.sort();
    if files.is_empty() {
        return Err(format!("no macro-enabled Office files found under {input}"));
    }

    println!("file\tmodules\tprocedures\tstatements\tunparsed\tclass\tformula-engine");
    let mut reports = Vec::new();
    let mut failures = Vec::new();
    for path in files {
        match inspect_project(&path) {
            Ok(report) => {
                let (procedures, statements, unparsed) = project_totals(&report);
                let class = project_class(&report).map_or("-", Class::as_str);
                let formula = if report
                    .modules
                    .iter()
                    .any(|module| module.analysis.needs_formula_engine)
                {
                    "yes"
                } else {
                    "no"
                };
                println!(
                    "{}\t{}\t{procedures}\t{statements}\t{unparsed}\t{class}\t{formula}",
                    report.path.display(),
                    report.modules.len()
                );
                reports.push(report);
            }
            Err(error) => {
                println!("{}\tERROR\t-\t-\t-\t-\t-", path.display());
                failures.push(format!("{}: {error}", path.display()));
            }
        }
    }

    print_duplicate_modules(&reports);
    print_related_modules(&reports);
    println!();
    println!(
        "Inventory: {} succeeded, {} failed",
        reports.len(),
        failures.len()
    );
    if failures.is_empty() {
        Ok(())
    } else {
        Err(failures.join("; "))
    }
}

pub(crate) fn inventory_json(input: &str) -> Result<(), String> {
    let root = Path::new(input);
    let mut files = Vec::new();
    collect_macro_files(root, &mut files)?;
    files.sort();
    if files.is_empty() {
        return Err(format!("no macro-enabled Office files found under {input}"));
    }

    let mut reports = Vec::new();
    let mut failures = Vec::new();
    for path in files {
        match inspect_project(&path) {
            Ok(report) => reports.push(report),
            Err(error) => failures.push((path, error)),
        }
    }

    let projects: Vec<_> = reports
        .iter()
        .map(|report| {
            let (procedures, statements, unparsed) = project_totals(report);
            let modules: Vec<_> = report
                .modules
                .iter()
                .map(|module| {
                    let findings: Vec<_> = module
                        .analysis
                        .findings
                        .iter()
                        .map(|finding| {
                            json!({
                                "line": finding.line,
                                "class": finding.class.map(Class::as_str),
                                "what": finding.what,
                                "reason": finding.reason,
                            })
                        })
                        .collect();
                    json!({
                        "name": module.name,
                        "class": module.analysis.class.map(Class::as_str),
                        "verdict": module.analysis.verdict(),
                        "metrics": {
                            "procedures": module.analysis.metrics.procedures,
                            "statements": module.analysis.metrics.statements,
                            "max_nesting": module.analysis.metrics.max_nesting,
                            "longest_procedure": module.analysis.metrics.longest_procedure,
                            "unparsed": module.analysis.metrics.unparsed,
                        },
                        "needs_formula_engine": module.analysis.needs_formula_engine,
                        "has_option_explicit": module.analysis.has_option_explicit,
                        "blanket_error_handlers": module.analysis.blanket_error_handlers,
                        "external_declares": module.analysis.external_declares,
                        "uncalled_procedures": module.analysis.uncalled_procedures,
                        "api_names": module.analysis.api_names,
                        "standard_fingerprint": format!("{:032x}", module.fingerprint.combined),
                        "findings": findings,
                    })
                })
                .collect();
            json!({
                "path": report.path.to_string_lossy(),
                "container_part": report.container_part,
                "class": project_class(report).map(Class::as_str),
                "needs_formula_engine": report.modules.iter().any(|module| module.analysis.needs_formula_engine),
                "metrics": {
                    "modules": report.modules.len(),
                    "procedures": procedures,
                    "statements": statements,
                    "unparsed": unparsed,
                },
                "modules": modules,
            })
        })
        .collect();
    let duplicate_groups = duplicate_module_groups(&reports);
    let related: Vec<_> = related_module_pairs(&reports)
        .into_iter()
        .map(|pair| {
            json!({
                "left": pair.left,
                "right": pair.right,
                "shared": pair.similarity.shared,
                "only_left": pair.similarity.only_a,
                "only_right": pair.similarity.only_b,
                "jaccard": pair.similarity.jaccard,
                "diverged": pair.similarity.diverged,
                "declarations_differ": pair.similarity.declarations_differ,
            })
        })
        .collect();
    let errors: Vec<_> = failures
        .iter()
        .map(|(path, error)| json!({ "path": path.to_string_lossy(), "error": error }))
        .collect();
    let output = json!({
        "schema": "oxivba-inventory-v1",
        "root": root.to_string_lossy(),
        "projects": projects,
        "duplicate_groups": duplicate_groups,
        "related_modules": related,
        "errors": errors,
    });
    println!(
        "{}",
        serde_json::to_string_pretty(&output)
            .map_err(|error| format!("cannot encode inventory JSON: {error}"))?
    );

    if failures.is_empty() {
        Ok(())
    } else {
        Err(format!("{} file(s) could not be analyzed", failures.len()))
    }
}

fn inspect_project(path: &Path) -> Result<ProjectReport, String> {
    let data =
        fs::read(path).map_err(|error| format!("cannot read {}: {error}", path.display()))?;
    let mut archive = OoxmlArchive::new(&data)
        .map_err(|error| format!("cannot open the OOXML package: {error}"))?;
    let container_part = archive
        .file_names()
        .into_iter()
        .find(|name| {
            name.rsplit('/')
                .next()
                .is_some_and(|leaf| leaf.eq_ignore_ascii_case("vbaProject.bin"))
        })
        .ok_or_else(|| "the package does not contain vbaProject.bin".to_string())?;
    let project_data = archive
        .read_binary_part(&container_part)
        .map_err(|error| format!("cannot read {container_part}: {error}"))?;
    let project = ovba::open_project(project_data)
        .map_err(|error| format!("cannot parse {container_part}: {error}"))?;

    let mut modules = Vec::with_capacity(project.modules.len());
    for module_info in &project.modules {
        let source = project
            .module_source(&module_info.name)
            .map_err(|error| format!("cannot extract module {}: {error}", module_info.name))?;
        let module = parse_module(&source)
            .map_err(|error| format!("cannot tokenize module {}: {error}", module_info.name))?;
        modules.push(ModuleReport {
            name: module_info.name.clone(),
            analysis: analyse(&module),
            fingerprint: fingerprint_module(&module, Strength::Standard),
        });
    }

    Ok(ProjectReport {
        path: path.to_path_buf(),
        container_part,
        modules,
    })
}

fn project_totals(report: &ProjectReport) -> (usize, usize, usize) {
    report.modules.iter().fold((0, 0, 0), |totals, module| {
        (
            totals.0 + module.analysis.metrics.procedures,
            totals.1 + module.analysis.metrics.statements,
            totals.2 + module.analysis.metrics.unparsed,
        )
    })
}

fn project_class(report: &ProjectReport) -> Option<Class> {
    report
        .modules
        .iter()
        .filter_map(|module| module.analysis.class)
        .max_by_key(|class| class.severity())
}

fn print_duplicate_modules(reports: &[ProjectReport]) {
    let duplicates = duplicate_module_groups(reports);
    if duplicates.is_empty() {
        return;
    }

    println!();
    println!("Structurally identical modules (standard fingerprint):");
    for members in duplicates {
        println!("  {}", members.join(" = "));
    }
}

fn duplicate_module_groups(reports: &[ProjectReport]) -> Vec<Vec<String>> {
    let mut groups: BTreeMap<u128, Vec<String>> = BTreeMap::new();
    for report in reports {
        for module in &report.modules {
            if module.fingerprint.procedures.is_empty() && module.fingerprint.declarations == 0 {
                continue;
            }
            groups
                .entry(module.fingerprint.combined)
                .or_default()
                .push(format!("{}::{}", report.path.display(), module.name));
        }
    }

    groups
        .into_values()
        .filter(|members| members.len() > 1)
        .collect()
}

fn print_related_modules(reports: &[ProjectReport]) {
    let related = related_module_pairs(reports);
    if related.is_empty() {
        return;
    }

    println!();
    println!("Related modules (standard fingerprint):");
    for pair in related {
        let diverged = if pair.similarity.diverged.is_empty() {
            String::new()
        } else {
            format!("; diverged: {}", pair.similarity.diverged.join(", "))
        };
        let declarations = if pair.similarity.declarations_differ {
            "; declarations differ"
        } else {
            ""
        };
        println!(
            "  {} <> {}: {:.1}% (shared {}; only {}/{}{diverged}{declarations})",
            pair.left,
            pair.right,
            pair.similarity.jaccard * 100.0,
            pair.similarity.shared,
            pair.similarity.only_a,
            pair.similarity.only_b
        );
    }
}

fn related_module_pairs(reports: &[ProjectReport]) -> Vec<RelatedModules> {
    let modules: Vec<_> = reports
        .iter()
        .flat_map(|report| {
            report.modules.iter().filter_map(move |module| {
                (!module.fingerprint.procedures.is_empty()).then_some((report, module))
            })
        })
        .collect();
    let mut by_hash: BTreeMap<u128, Vec<usize>> = BTreeMap::new();
    let mut by_name: BTreeMap<String, Vec<usize>> = BTreeMap::new();
    for (index, (_, module)) in modules.iter().enumerate() {
        for procedure in &module.fingerprint.procedures {
            by_hash.entry(procedure.hash).or_default().push(index);
            by_name
                .entry(procedure.name.to_ascii_lowercase())
                .or_default()
                .push(index);
        }
    }

    let mut candidates = BTreeSet::new();
    for members in by_hash.values().chain(by_name.values()) {
        for (position, &left) in members.iter().enumerate() {
            for &right in &members[position + 1..] {
                if left != right {
                    candidates.insert((left.min(right), left.max(right)));
                }
            }
        }
    }

    let mut related = Vec::new();
    for (left, right) in candidates {
        let (left_report, left_module) = modules[left];
        let (right_report, right_module) = modules[right];
        if left_report.path == right_report.path
            || left_module.fingerprint.combined == right_module.fingerprint.combined
        {
            continue;
        }
        let similarity = compare(&left_module.fingerprint, &right_module.fingerprint);
        if (similarity.shared > 0 && similarity.jaccard >= 0.5) || !similarity.diverged.is_empty() {
            related.push(RelatedModules {
                left: format!("{}::{}", left_report.path.display(), left_module.name),
                right: format!("{}::{}", right_report.path.display(), right_module.name),
                similarity,
            });
        }
    }
    related.sort_by(|a, b| {
        b.similarity
            .jaccard
            .total_cmp(&a.similarity.jaccard)
            .then_with(|| a.left.cmp(&b.left))
            .then_with(|| a.right.cmp(&b.right))
    });
    related
}

fn collect_macro_files(path: &Path, files: &mut Vec<PathBuf>) -> Result<(), String> {
    if path.is_file() {
        if is_macro_office_file(path) {
            files.push(path.to_path_buf());
        }
        return Ok(());
    }
    if !path.is_dir() {
        return Err(format!("{} is not a file or directory", path.display()));
    }

    let entries = fs::read_dir(path)
        .map_err(|error| format!("cannot read directory {}: {error}", path.display()))?;
    for entry in entries {
        let entry = entry.map_err(|error| {
            format!(
                "cannot read a directory entry under {}: {error}",
                path.display()
            )
        })?;
        let child = entry.path();
        if child.is_dir() {
            collect_macro_files(&child, files)?;
        } else if is_macro_office_file(&child) {
            files.push(child);
        }
    }
    Ok(())
}

fn is_macro_office_file(path: &Path) -> bool {
    path.extension()
        .and_then(|extension| extension.to_str())
        .is_some_and(|extension| {
            matches!(
                extension.to_ascii_lowercase().as_str(),
                "xlsm" | "xlam" | "xlsb" | "docm" | "dotm" | "pptm" | "potm" | "ppam"
            )
        })
}
