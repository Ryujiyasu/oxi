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
use oxihanko::attest::{clearance, digest_project, Attestation, Clearance};
use oxihanko::verify::TrustedSigners;
use ring::rand::SystemRandom;
use ring::signature::{Ed25519KeyPair, KeyPair};
use oxihanko::attest::SignatureVerifier;
use oxivba_core::safety::{assess, Capability, SafetyReport};
use oxivba_core::{analyse, parse_module, Analysis, Class};
use serde_json::json;

struct ModuleReport {
    name: String,
    source: String,
    analysis: Analysis,
    fingerprint: ModuleFingerprint,
    safety: SafetyReport,
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

/// Says what the macros in a file would be able to reach, without opening the
/// file in Office and without running anything.
///
/// Deliberately gives no verdict beyond one: whether opening the file is by
/// itself enough to run code. `Shell` is how a legitimate macro opens a PDF,
/// so a tool that calls it dangerous teaches its reader to click through.
///
/// Exit status is the material, not a judgement: `0` when there is nothing for
/// a person to read, `3` when there is.
pub(crate) fn safety(input: &str) -> Result<bool, String> {
    let report = inspect_project(Path::new(input))?;

    println!("VBA project: {}", report.path.display());
    println!("Container part: {}", report.container_part);
    println!("Modules: {}", report.modules.len());
    println!();

    let sealed = read_attestation(Path::new(input))?;
    let digest = digest_project(
        report
            .modules
            .iter()
            .map(|module| (module.name.as_str(), module.source.as_str())),
    );
    let trusted = trusted_signers(Path::new(input))?;
    if sealed.is_some() && trusted.is_none() {
        println!("There is a seal beside this file, but nobody has been named as");
        println!("trusted to place one. Put the signers in oxi-hanko-signers.txt");
        println!("beside the workbook, or point OXI_HANKO_SIGNERS at the list.");
        println!();
    }
    match clearance(
        &digest,
        sealed.as_ref(),
        trusted.as_ref().map(|signers| signers as &dyn SignatureVerifier),
    ) {
        Clearance::Cleared { signer, note } => {
            println!("Sealed by {signer}: {note}");
            println!("Nothing to read: this exact code has been cleared.");
            return Ok(false);
        }
        Clearance::MatchesButUnsigned => {
            println!("A seal beside this file matches the code, but NOTHING SIGNED IT.");
            println!("Anyone who can edit the workbook can write that file, so it");
            println!("clears nothing. Read on.");
            println!();
        }
        Clearance::MatchesButNobodyChecked => {
            println!("A seal beside this file matches the code and is signed, but");
            println!("nobody here is trusted to place one, so whose signature it is");
            println!("has not been established. Read on.");
            println!();
        }
        Clearance::SignatureRejected { why } => {
            println!("A seal beside this file did NOT verify: {why}");
            println!();
        }
        Clearance::Stale(changed) => {
            println!("A seal beside this file is over OLDER code. What moved since:");
            for name in &changed.edited {
                println!("  edited  {name}");
            }
            for name in &changed.added {
                println!("  added   {name}");
            }
            for name in &changed.removed {
                println!("  removed {name}");
            }
            println!("Only these need reading again.");
            println!();
        }
        Clearance::NotSealed => {}
    }

    let mut whole = SafetyReport::default();
    let mut opens: Vec<(&str, &str)> = Vec::new();
    for module in &report.modules {
        for name in &module.safety.runs_without_asking {
            opens.push((module.name.as_str(), name.as_str()));
        }
        whole.signals.extend(module.safety.signals.iter().cloned());
        whole.unread_lines += module.safety.unread_lines;
        whole
            .unresolved_late_binding
            .extend(module.safety.unresolved_late_binding.iter().copied());
    }

    if opens.is_empty() {
        println!("Opening the file does not by itself run anything.");
    } else {
        println!("RUNS WHEN THE FILE IS OPENED:");
        for (module, procedure) in &opens {
            println!("  {module}.{procedure}");
        }
        println!("  Opening this file is the trigger, so the question is not");
        println!("  whether to run the macros but whether they have already run.");
    }
    println!();

    for module in &report.modules {
        if !module.safety.needs_a_reader() {
            continue;
        }
        println!("[{}]", module.name);
        for signal in &module.safety.signals {
            let place = if signal.line == 0 {
                "     ".to_string()
            } else {
                format!("{:>5}", signal.line)
            };
            println!(
                "  {place}  {}: {} — {}",
                signal.capability.label(),
                signal.what,
                signal.reason
            );
        }
        if module.safety.unread_lines > 0 {
            println!(
                "         {} line(s) could not be read, so nobody has cleared them",
                module.safety.unread_lines
            );
        }
        println!();
    }

    let quiet = report
        .modules
        .iter()
        .filter(|module| !module.safety.needs_a_reader())
        .count();
    println!(
        "Summary: {} module(s), {quiet} with nothing to read.",
        report.modules.len()
    );
    let mut counts: Vec<(Capability, usize)> = whole.capabilities().into_iter().collect();
    counts.sort_by(|a, b| b.1.cmp(&a.1).then(a.0.cmp(&b.0)));
    for (capability, count) in counts {
        println!("  {count:>5}  {}", capability.label());
    }
    if !whole.unresolved_late_binding.is_empty() {
        println!(
            "  {:>5}  CreateObject/GetObject whose target is worked out while it runs",
            whole.unresolved_late_binding.len()
        );
    }
    Ok(whole.needs_a_reader() || !opens.is_empty())
}

/// The seal that belongs to a workbook, if one is beside it: the same path
/// with `.hanko.json` appended, so it travels with the file.
fn read_attestation(workbook: &Path) -> Result<Option<Attestation>, String> {
    let mut name = workbook.as_os_str().to_os_string();
    name.push(".hanko.json");
    let beside = PathBuf::from(name);
    if !beside.is_file() {
        return Ok(None);
    }
    let text = fs::read_to_string(&beside)
        .map_err(|error| format!("cannot read {}: {error}", beside.display()))?;
    serde_json::from_str(&text)
        .map(Some)
        .map_err(|error| format!("cannot read the seal {}: {error}", beside.display()))
}

/// Writes a new Ed25519 signing key, and prints the public half to put in
/// the trusted-signers list.
///
/// The private key is written with no passphrase, because a passphrase this
/// tool invents would be a passphrase nobody keeps. Whoever holds this file
/// can seal macros in their name.
pub(crate) fn keygen(path: &str) -> Result<(), String> {
    if Path::new(path).exists() {
        return Err(format!("{path} already exists; refusing to overwrite a key"));
    }
    let random = SystemRandom::new();
    let document = Ed25519KeyPair::generate_pkcs8(&random)
        .map_err(|_| "cannot generate a key".to_string())?;
    let key = Ed25519KeyPair::from_pkcs8(document.as_ref())
        .map_err(|_| "the generated key did not read back".to_string())?;
    fs::write(path, document.as_ref())
        .map_err(|error| format!("cannot write {path}: {error}"))?;
    println!("Wrote {path}. Keep it; anyone holding it can seal in your name.");
    println!();
    println!("Add this line to oxi-hanko-signers.txt, with your own name:");
    println!("<your name>\t{}", key_fingerprint(&key));
    Ok(())
}

/// The people trusted to seal macros, from `OXI_HANKO_SIGNERS` or from
/// `oxi-hanko-signers.txt` beside the workbook.
///
/// Returning `None` is not a failure: it means nobody has been named yet, and
/// the report says so rather than quietly clearing nothing.
fn trusted_signers(workbook: &Path) -> Result<Option<TrustedSigners>, String> {
    let listed = match std::env::var("OXI_HANKO_SIGNERS") {
        Ok(path) => PathBuf::from(path),
        Err(_) => {
            let beside = workbook
                .parent()
                .unwrap_or(Path::new("."))
                .join("oxi-hanko-signers.txt");
            if !beside.is_file() {
                return Ok(None);
            }
            beside
        }
    };
    let text = fs::read_to_string(&listed)
        .map_err(|error| format!("cannot read {}: {error}", listed.display()))?;
    let signers = TrustedSigners::from_list(&text)
        .map_err(|error| format!("{}: {error}", listed.display()))?;
    Ok(Some(signers))
}

/// Writes a seal beside a workbook, signed with an Ed25519 key in PKCS#8.
///
/// The signer's NAME is not taken from here; it comes from whichever trusted
/// key verifies at read time. What is written down is what was read and when,
/// in the sealer's words.
pub(crate) fn seal(input: &str, key_path: &str, note: &str) -> Result<(), String> {
    let report = inspect_project(Path::new(input))?;
    let digest = digest_project(
        report
            .modules
            .iter()
            .map(|module| (module.name.as_str(), module.source.as_str())),
    );
    let key_bytes = fs::read(key_path)
        .map_err(|error| format!("cannot read the key {key_path}: {error}"))?;
    // PKCS#8 v2 carries the public key beside the private one and v1 does
    // not. `ring` generates v2 and checks it; OpenSSL and Python's
    // cryptography write v1. Refusing v1 would mean refusing most keys
    // anyone already has, so both are read.
    let key = Ed25519KeyPair::from_pkcs8(&key_bytes)
        .or_else(|_| Ed25519KeyPair::from_pkcs8_maybe_unchecked(&key_bytes))
        .map_err(|_| format!("{key_path} is not an Ed25519 private key in PKCS#8"))?;

    let mut attestation = Attestation {
        digest,
        signer: String::new(),
        sealed_at: today(),
        note: note.to_string(),
        signature: None,
        certificate: None,
    };
    attestation.signer = key_fingerprint(&key);
    let signature = key.sign(&attestation.signed_bytes());
    attestation.signature = Some(signature.as_ref().to_vec());

    let mut name = Path::new(input).as_os_str().to_os_string();
    name.push(".hanko.json");
    let beside = PathBuf::from(name);
    let text = serde_json::to_string_pretty(&attestation)
        .map_err(|error| format!("cannot write the seal: {error}"))?;
    fs::write(&beside, text)
        .map_err(|error| format!("cannot write {}: {error}", beside.display()))?;
    println!("Sealed {} module(s) as {}", report.modules.len(), beside.display());
    println!("Project digest: {}", attestation.digest.project);
    println!();
    println!("This seal clears the file only for readers who trust the key that");
    println!("signed it. Add its public key to oxi-hanko-signers.txt as");
    println!("  <name><TAB>{}", key_fingerprint(&key));
    Ok(())
}

fn key_fingerprint(key: &Ed25519KeyPair) -> String {
    key.public_key()
        .as_ref()
        .iter()
        .map(|byte| format!("{byte:02x}"))
        .collect()
}

/// The date, from the system clock, as `YYYY-MM-DD`. Written by whoever seals
/// and never checked when reading, so it is a note to a human rather than a
/// fact the seal rests on.
fn today() -> String {
    let now = std::time::SystemTime::now()
        .duration_since(std::time::UNIX_EPOCH)
        .map(|since| since.as_secs() as i64)
        .unwrap_or(0);
    let days = now.div_euclid(86_400);
    let (year, month, day) = civil_from_days(days);
    format!("{year:04}-{month:02}-{day:02}")
}

/// Howard Hinnant's days-to-civil-date, which needs no calendar crate.
fn civil_from_days(days: i64) -> (i64, u32, u32) {
    let z = days + 719_468;
    let era = z.div_euclid(146_097);
    let doe = z.rem_euclid(146_097);
    let yoe = (doe - doe / 1_460 + doe / 36_524 - doe / 146_096) / 365;
    let y = yoe + era * 400;
    let doy = doe - (365 * yoe + yoe / 4 - yoe / 100);
    let mp = (5 * doy + 2) / 153;
    let d = (doy - (153 * mp + 2) / 5 + 1) as u32;
    let m = if mp < 10 { mp + 3 } else { mp - 9 } as u32;
    (if m <= 2 { y + 1 } else { y }, m, d)
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
        let analysis = analyse(&module);
        let safety = assess(&module, &analysis);
        modules.push(ModuleReport {
            name: module_info.name.clone(),
            source,
            analysis,
            fingerprint: fingerprint_module(&module, Strength::Standard),
            safety,
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
