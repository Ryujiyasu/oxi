// SPDX-License-Identifier: MIT OR Apache-2.0

//! Reading a workbook's macros in the browser, before anyone runs them.
//!
//! This is the same read the CLI does, at the point where it is most useful:
//! the moment someone attaches an `.xlsm`, with nothing executed and Excel
//! nowhere in sight. The whole path is pure Rust — unzip the package, walk the
//! `vbaProject.bin` compound file, decompress each module, parse it — so it
//! runs in a tab as readily as on a server.
//!
//! What comes back is evidence, not a verdict. See `oxivba_core::safety` for
//! why: a report that shouts at every workbook is a report nobody reads.
//!
//! The project digest travels with it so a caller can ask, somewhere that
//! holds the trusted keys, whether this exact code has been sealed. Verifying
//! a seal is deliberately NOT done here — that needs a signature stack, and a
//! browser build has no business carrying one.

use oxidocs_common::archive::OoxmlArchive;
use oxivba_core::safety::assess;
use oxivba_core::{analyse, parse_module};
use serde::Serialize;
use wasm_bindgen::prelude::*;

#[derive(Serialize)]
pub struct MacroSafetyReport {
    /// The part the macros were found in, normally `xl/vbaProject.bin`.
    container_part: String,
    /// SHA-256 over every module's name and source. The key a seal is placed
    /// on; see `oxihanko::attest`.
    project_digest: String,
    modules: Vec<ModuleSafety>,
    /// Procedures Office runs on its own, as `Module.Procedure`. Non-empty
    /// means opening the file is by itself enough to run code.
    runs_without_asking: Vec<String>,
    /// `(what it can reach, how much evidence there is)`, most first.
    capabilities: Vec<CapabilityCount>,
    /// How many `CreateObject`/`GetObject` calls work out their target while
    /// they run, so reading cannot say what they reach.
    unresolved_late_binding: usize,
    /// Lines no parser could read. Not the same as safe.
    unread_lines: usize,
    /// Whether a person still has to read something.
    needs_a_reader: bool,
}

#[derive(Serialize)]
struct ModuleSafety {
    name: String,
    /// The module's own SHA-256, so a caller can show which one moved.
    digest: String,
    signals: Vec<SafetySignal>,
    unread_lines: usize,
}

#[derive(Serialize)]
struct SafetySignal {
    capability: String,
    what: String,
    reason: String,
    /// `0` where the evidence is a name counted across the module rather than
    /// one statement.
    line: u32,
}

#[derive(Serialize)]
struct CapabilityCount {
    capability: String,
    evidence: usize,
}

/// Reads the macros in an `.xlsm` / `.xlam` / `.docm` and says what they could
/// reach. Nothing is executed.
///
/// Returns an error only when the bytes are not an Office package. A package
/// with no macros in it is not an error: it answers with an empty report,
/// because "there are no macros" is exactly what a caller wants to hear.
#[wasm_bindgen]
pub fn read_macro_safety(package: &[u8]) -> Result<JsValue, JsError> {
    let report = read(package).map_err(|error| JsError::new(&error))?;
    serde_wasm_bindgen::to_value(&report).map_err(|error| JsError::new(&error.to_string()))
}

fn read(package: &[u8]) -> Result<MacroSafetyReport, String> {
    let mut archive =
        OoxmlArchive::new(package).map_err(|error| format!("cannot open the package: {error}"))?;
    let Some(container_part) = archive.file_names().into_iter().find(|name| {
        name.rsplit('/')
            .next()
            .is_some_and(|leaf| leaf.eq_ignore_ascii_case("vbaProject.bin"))
    }) else {
        return Ok(MacroSafetyReport {
            container_part: String::new(),
            project_digest: oxihanko::attest::digest_project([]).project,
            modules: Vec::new(),
            runs_without_asking: Vec::new(),
            capabilities: Vec::new(),
            unresolved_late_binding: 0,
            unread_lines: 0,
            needs_a_reader: false,
        });
    };

    let container = archive
        .read_binary_part(&container_part)
        .map_err(|error| format!("cannot read {container_part}: {error}"))?;
    let project = ovba::open_project(container)
        .map_err(|error| format!("cannot read the VBA project: {error}"))?;

    let mut sources = Vec::with_capacity(project.modules.len());
    for module in &project.modules {
        let source = project
            .module_source(&module.name)
            .map_err(|error| format!("cannot read module {}: {error}", module.name))?;
        sources.push((module.name.clone(), source));
    }

    let digest = oxihanko::attest::digest_project(
        sources
            .iter()
            .map(|(name, source)| (name.as_str(), source.as_str())),
    );

    let mut modules = Vec::with_capacity(sources.len());
    let mut runs_without_asking = Vec::new();
    let mut counts: std::collections::BTreeMap<String, usize> = std::collections::BTreeMap::new();
    let mut unresolved = 0usize;
    let mut unread = 0usize;
    let mut needs_a_reader = false;

    for (name, source) in &sources {
        // A module the parser cannot even tokenise is reported as such rather
        // than skipped: silence here would read as "nothing to worry about".
        let Ok(parsed) = parse_module(source) else {
            needs_a_reader = true;
            modules.push(ModuleSafety {
                name: name.clone(),
                digest: digest
                    .modules
                    .iter()
                    .find(|(module, _)| module == name)
                    .map(|(_, hash)| hash.clone())
                    .unwrap_or_default(),
                signals: vec![SafetySignal {
                    capability: "could not be read".to_string(),
                    what: name.clone(),
                    reason: "this module did not parse, so nothing about it has been established"
                        .to_string(),
                    line: 0,
                }],
                unread_lines: source.lines().count(),
            });
            unread += source.lines().count();
            continue;
        };
        let analysis = analyse(&parsed);
        let safety = assess(&parsed, &analysis);

        for procedure in &safety.runs_without_asking {
            runs_without_asking.push(format!("{name}.{procedure}"));
        }
        for (capability, count) in safety.capabilities() {
            *counts.entry(capability.label().to_string()).or_default() += count;
        }
        unresolved += safety.unresolved_late_binding.len();
        unread += safety.unread_lines;
        needs_a_reader |= safety.needs_a_reader();

        modules.push(ModuleSafety {
            name: name.clone(),
            digest: digest
                .modules
                .iter()
                .find(|(module, _)| module == name)
                .map(|(_, hash)| hash.clone())
                .unwrap_or_default(),
            signals: safety
                .signals
                .iter()
                .map(|signal| SafetySignal {
                    capability: signal.capability.label().to_string(),
                    what: signal.what.clone(),
                    reason: signal.reason.clone(),
                    line: signal.line,
                })
                .collect(),
            unread_lines: safety.unread_lines,
        });
    }

    let mut capabilities: Vec<CapabilityCount> = counts
        .into_iter()
        .map(|(capability, evidence)| CapabilityCount {
            capability,
            evidence,
        })
        .collect();
    capabilities.sort_by(|a, b| {
        b.evidence
            .cmp(&a.evidence)
            .then(a.capability.cmp(&b.capability))
    });

    Ok(MacroSafetyReport {
        container_part,
        project_digest: digest.project,
        modules,
        runs_without_asking,
        capabilities,
        unresolved_late_binding: unresolved,
        unread_lines: unread,
        needs_a_reader,
    })
}

/// The same read, as JSON, for callers that are not a browser: a server that
/// screens an upload before it reaches anyone, or a test.
pub fn read_macro_safety_native(package: &[u8]) -> Result<String, String> {
    let report = read(package)?;
    serde_json::to_string_pretty(&report).map_err(|error| error.to_string())
}

#[cfg(test)]
mod tests {
    use super::*;

    /// Bytes that are not an Office package are an error; a package with no
    /// macros is not.
    #[test]
    fn something_that_is_not_a_package_is_an_error() {
        assert!(read(b"not a zip at all").is_err());
    }

    #[test]
    fn a_package_without_macros_answers_quietly() {
        // The smallest thing OoxmlArchive will open: an empty zip.
        let empty_zip = [
            0x50, 0x4b, 0x05, 0x06, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0,
        ];
        match read(&empty_zip) {
            Ok(report) => {
                assert!(!report.needs_a_reader);
                assert!(report.modules.is_empty());
                assert!(report.container_part.is_empty());
            }
            Err(error) => panic!("an empty package should read quietly: {error}"),
        }
    }
}
