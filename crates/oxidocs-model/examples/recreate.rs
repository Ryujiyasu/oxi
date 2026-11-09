// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! Make every document in a corpus from nothing with edits, and say which
//! ones come back the same.
//!
//!     cargo run -p oxidocs-model --release --example recreate -- \
//!         [--keep DIR] [corpus_dir] [report.json]
//!
//! For each .docx: read it into the model, write the edits that build it
//! (`script`), pass them through JSON, apply them to an empty document,
//! write that, and compare it with the original canonically (runs joined:
//! where a run of the same formatting is cut says nothing). The report
//! names, per document, what no edit makes yet or where the result differs.

use std::collections::BTreeMap;
use std::path::{Path, PathBuf};

use oxidocs_model::edit::{apply_all, Edit};
use oxidocs_model::{canon, script, DocxModel, Package};

fn docs(dir: &Path, out: &mut Vec<PathBuf>) {
    let Ok(rd) = std::fs::read_dir(dir) else { return };
    for e in rd.flatten() {
        let p = e.path();
        if p.is_dir() {
            docs(&p, out);
        } else if p.extension().map(|x| x == "docx").unwrap_or(false)
            && !p.file_name().unwrap().to_string_lossy().starts_with("~$")
        {
            out.push(p);
        }
    }
}

/// A path without element indexes, keeping the last two steps.
fn feature(path: &str) -> String {
    let mut out = String::new();
    let mut skip = false;
    for c in path.chars() {
        match c {
            '[' => skip = true,
            ']' => skip = false,
            c if !skip => out.push(c),
            _ => {}
        }
    }
    let steps: Vec<&str> = out.split('/').filter(|s| !s.is_empty()).collect();
    steps[steps.len().saturating_sub(2)..].join("/")
}

/// (detail, cause) for one document; the file made, when one was.
fn recreate(bytes: &[u8]) -> (String, String, Option<Vec<u8>>) {
    let mut made = None;
    let (detail, cause) = recreate_into(bytes, &mut made);
    (detail, cause, made)
}

fn recreate_into(bytes: &[u8], kept: &mut Option<Vec<u8>>) -> (String, String) {
    let Ok(pkg) = Package::read(bytes) else { return ("unreadable".into(), "unreadable".into()) };
    let Ok(model) = DocxModel::from_package(&pkg) else { return ("model refuses".into(), "model refuses".into()) };
    let edits = match script::script(&model) {
        Ok(e) => e,
        Err(m) => {
            let cause = m.0.split(':').next().unwrap_or("").to_string();
            return (format!("missing {}", m.0), format!("missing {cause}"));
        }
    };
    let json = serde_json::to_string(&edits).unwrap();
    let edits: Vec<Edit> = match oxidocs_model::edit::from_json(&json) {
        Ok(e) => e,
        Err(e) => return (format!("json {e}"), "json".into()),
    };
    let mut made = DocxModel::default();
    if let Err(e) = apply_all(&mut made, &edits) {
        return (format!("replay {e}"), "replay".into());
    }
    let out = match made.write(None) {
        Ok(o) => o,
        Err(e) => return (format!("write {e}"), format!("write {}", e.to_string().chars().take(50).collect::<String>())),
    };
    let b = Package::read(&out).unwrap();
    *kept = Some(out);
    let diffs = canon::compare_content(&pkg, &b);
    if diffs.is_empty() {
        return (format!("ok {}", edits.len()), "ok".into());
    }
    let d = &diffs[0];
    let all: Vec<String> = diffs.iter().take(3).map(|d| d.to_string()).collect();
    (format!("differs {}", all.join(" | ")), format!("differs {} {}", d.part, feature(&d.path)))
}

fn main() {
    let mut args: Vec<String> = std::env::args().skip(1).collect();
    // --keep DIR writes every file made there (to hand to Word).
    let keep = args.iter().position(|a| a == "--keep").map(|i| {
        let dir = PathBuf::from(args.get(i + 1).cloned().expect("--keep DIR"));
        args.drain(i..=i + 1);
        std::fs::create_dir_all(&dir).unwrap();
        dir
    });
    let corpus = PathBuf::from(args.first().cloned().unwrap_or_else(|| "tools/golden-test/documents/docx".into()));
    let report = PathBuf::from(args.get(1).cloned().unwrap_or_else(|| "pipeline_data/edit_coverage/recreate.json".into()));
    let mut files = Vec::new();
    docs(&corpus, &mut files);
    files.sort();

    let mut per_doc = BTreeMap::new();
    let mut ok = 0;
    let mut causes: BTreeMap<String, Vec<String>> = BTreeMap::new();
    let started = std::time::Instant::now();
    for (k, f) in files.iter().enumerate() {
        let name = f.strip_prefix(&corpus).unwrap_or(f).to_string_lossy().replace('\\', "/");
        let bytes = std::fs::read(f).unwrap();
        let t = std::time::Instant::now();
        let (detail, cause, out) = std::panic::catch_unwind(|| recreate(&bytes)).unwrap_or_else(|_| ("panic".into(), "panic".into(), None));
        if let (Some(dir), Some(out)) = (&keep, out) {
            std::fs::write(dir.join(f.file_name().unwrap()), out).unwrap();
        }
        let ms = t.elapsed().as_millis();
        if ms > 20_000 {
            eprintln!("slow {ms}ms {name}");
        }
        if cause == "ok" {
            ok += 1;
        } else {
            causes.entry(cause).or_default().push(name.clone());
        }
        per_doc.insert(name, serde_json::json!({ "verdict": detail, "ms": ms }));
        if (k + 1) % 500 == 0 {
            eprintln!("{} / {} ({} ok, {:.0}s)", k + 1, files.len(), ok, started.elapsed().as_secs_f64());
        }
    }
    let mut ranked: Vec<(&String, &Vec<String>)> = causes.iter().collect();
    ranked.sort_by_key(|(_, v)| std::cmp::Reverse(v.len()));
    println!("recreated {}/{}", ok, files.len());
    for (cause, docs) in ranked.iter().take(40) {
        println!("  {:5}  {}   e.g. {}", docs.len(), cause, docs[0]);
    }
    if let Some(dir) = report.parent() {
        let _ = std::fs::create_dir_all(dir);
    }
    let summary = serde_json::json!({ "docs": files.len(), "recreated": ok, "causes": causes, "per_doc": per_doc });
    std::fs::write(&report, serde_json::to_string_pretty(&summary).unwrap()).unwrap();
}
