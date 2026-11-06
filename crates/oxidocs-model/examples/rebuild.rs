// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! Rebuild every document in a corpus through the model and say which ones
//! come back the same.
//!
//!     cargo run -p oxidocs-model --release --example rebuild -- \
//!         [--blank] [corpus_dir] [report.json]
//!
//! For each .docx: read it into the model, pass the model through JSON, write
//! it again (over the original as template, or from nothing with --blank),
//! and compare the two packages canonically. A document is rebuilt when the
//! comparison finds no difference. Otherwise the report names the first
//! feature the model refused, or the first place the output differs.

use std::collections::BTreeMap;
use std::path::{Path, PathBuf};

use oxidocs_model::{canon, rebuild, Error, Package};

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

fn is_real(name: &str) -> bool {
    let synthetic = ["gen2_", "gen_", "repro_", "test_", "pixel_", "sweep_", "indent_", "prog_", "d7_", "d11_", "_pixel_tmp"];
    !(synthetic.iter().any(|s| name.starts_with(s)) || name.starts_with("probe") || name.starts_with("prob"))
}

/// A feature path without element indexes, for counting across documents.
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
    // Keep the last two steps: enough to name the feature, not the place.
    let steps: Vec<&str> = out.split('/').filter(|s| !s.is_empty()).collect();
    steps[steps.len().saturating_sub(2)..].join("/")
}

/// Change one thing of the given kind somewhere in the model's JSON; say
/// whether anything was found to change.
fn sabotage(v: &mut serde_json::Value, kind: &str, in_hf: bool) -> bool {
    use serde_json::Value;
    match v {
        Value::Object(m) => {
            for (k, x) in m.iter_mut() {
                let here = in_hf || k == "headers_footers";
                match (kind, k.as_str(), &mut *x) {
                    ("text", "Text", Value::String(t)) if !in_hf => {
                        t.push('x');
                        return true;
                    }
                    ("header", "Text", Value::String(t)) if in_hf => {
                        t.push('x');
                        return true;
                    }
                    ("image", "data", Value::Array(bytes)) if !bytes.is_empty() => {
                        let b = bytes[0].as_u64().unwrap_or(0);
                        bytes[0] = Value::from((b + 1) % 256);
                        return true;
                    }
                    ("note", "footnotes", Value::Array(notes)) if notes.len() > 2 => {
                        notes.pop();
                        return true;
                    }
                    ("style", "styles", Value::Object(st)) => {
                        if let Some(Value::Array(list)) = st.get_mut("styles") {
                            if let Some(Value::Object(first)) = list.first_mut() {
                                first.insert("meta".into(), Value::Array(vec![]));
                                if let Some(Value::Array(a)) = first.get_mut("attrs") {
                                    a.push(serde_json::json!(["w:customStyle", "1"]));
                                }
                                return true;
                            }
                        }
                    }
                    ("theme", "theme", Value::Object(t)) => {
                        if let Some(Value::Object(i)) = t.get_mut("Internal") {
                            if let Some(Value::Array(bytes)) = i.get_mut("data") {
                                // Flip the case of a letter inside an attribute value
                                // (a typeface or colour name) in the middle of the part.
                                let mid = bytes.len() / 2;
                                let mut in_value = false;
                                for (k, slot) in bytes.iter_mut().enumerate() {
                                    let b = slot.as_u64().unwrap_or(0) as u8;
                                    if b == b'"' {
                                        in_value = !in_value;
                                    }
                                    if k > mid && in_value && b.is_ascii_alphabetic() {
                                        *slot = Value::from((b ^ 0x20) as u64);
                                        return true;
                                    }
                                }
                            }
                        }
                    }
                    ("prop", "attrs", Value::Array(a)) if !a.is_empty() && !in_hf => {
                        if let Some(Value::Array(pair)) = a.first_mut() {
                            if let Some(Value::String(val)) = pair.get_mut(1) {
                                val.push('9');
                                return true;
                            }
                        }
                    }
                    _ => {}
                }
                if sabotage(x, kind, here) {
                    return true;
                }
            }
            false
        }
        Value::Array(a) => a.iter_mut().any(|x| sabotage(x, kind, in_hf)),
        _ => false,
    }
}

fn main() {
    let mut args: Vec<String> = std::env::args().skip(1).collect();
    let blank = args.iter().any(|a| a == "--blank");
    let check = args.iter().any(|a| a == "--sabotage");
    // --keep DIR writes every rebuilt file there (to hand to Word).
    let keep = args.iter().position(|a| a == "--keep").map(|i| {
        let dir = PathBuf::from(args.get(i + 1).cloned().expect("--keep DIR"));
        args.drain(i..=i + 1);
        dir
    });
    if let Some(dir) = &keep {
        std::fs::create_dir_all(dir).unwrap();
    }
    args.retain(|a| a != "--blank" && a != "--sabotage");
    if check {
        let corpus = PathBuf::from(args.first().cloned().unwrap_or_else(|| "tools/golden-test/documents/docx".into()));
        let mut files = Vec::new();
        docs(&corpus, &mut files);
        files.sort();
        for kind in ["text", "header", "image", "note", "prop", "style", "theme"] {
            let (mut tried, mut caught) = (0, 0);
            let mut missed = Vec::new();
            for f in &files {
                let bytes = std::fs::read(f).unwrap();
                let pkg = Package::read(&bytes).unwrap();
                let Ok(model) = oxidocs_model::DocxModel::from_package(&pkg) else { continue };
                let mut v = serde_json::to_value(&model).unwrap();
                if !sabotage(&mut v, kind, false) {
                    continue;
                }
                let Ok(broken) = oxidocs_model::from_json(&v.to_string()) else { continue };
                let Ok(out) = broken.write((!blank).then_some(&pkg)) else { continue };
                tried += 1;
                if canon::compare(&pkg, &Package::read(&out).unwrap()).is_empty() {
                    missed.push(f.file_name().unwrap().to_string_lossy().to_string());
                } else {
                    caught += 1;
                }
            }
            println!("sabotage {kind:7}: caught {caught}/{tried}  missed {:?}", &missed[..missed.len().min(5)]);
        }
        return;
    }
    let corpus = PathBuf::from(args.first().cloned().unwrap_or_else(|| "tools/golden-test/documents/docx".into()));
    let report = PathBuf::from(args.get(1).cloned().unwrap_or_else(|| {
        format!("pipeline_data/edit_coverage/rebuild_{}.json", if blank { "blank" } else { "template" })
    }));
    let mut files = Vec::new();
    docs(&corpus, &mut files);
    files.sort();

    let mut per_doc = BTreeMap::new();
    let mut ok_all = 0;
    let mut ok_real = 0;
    let mut real = 0;
    let mut causes: BTreeMap<String, Vec<String>> = BTreeMap::new();
    for f in &files {
        let name = f.strip_prefix(&corpus).unwrap_or(f).to_string_lossy().replace('\\', "/");
        let bytes = std::fs::read(f).unwrap();
        let real_doc = is_real(f.file_name().unwrap().to_string_lossy().as_ref());
        real += real_doc as usize;
        let verdict = match std::panic::catch_unwind(|| rebuild(&bytes, !blank)) {
            Err(_) => ("panic".to_string(), "panic".to_string()),
            Ok(Err(Error::Unsupported(u))) => (format!("unsupported {}", u.path), format!("unsupported {}", feature(&u.path))),
            Ok(Err(e)) => (format!("error {e}"), format!("error {}", e.to_string().chars().take(60).collect::<String>())),
            Ok(Ok(out)) => {
                if let Some(dir) = &keep {
                    std::fs::write(dir.join(f.file_name().unwrap()), &out).unwrap();
                }
                let a = Package::read(&bytes).unwrap();
                let b = Package::read(&out).unwrap();
                let diffs = canon::compare(&a, &b);
                // Word refuses a part that names an undeclared prefix; only
                // count it when the original was clean.
                let ours: Vec<_> = canon::undeclared_prefixes(&b)
                    .into_iter()
                    .filter(|x| !canon::undeclared_prefixes(&a).contains(x))
                    .collect();
                if !ours.is_empty() {
                    (format!("undeclared prefix {ours:?}"), "undeclared prefix".to_string())
                } else if diffs.is_empty() {
                    ("ok".to_string(), "ok".to_string())
                } else {
                    let d = &diffs[0];
                    let all: Vec<String> = diffs.iter().take(3).map(|d| d.to_string()).collect();
                    (format!("differs {}", all.join(" | ")), format!("differs {} {}", d.part, feature(&d.path)))
                }
            }
        };
        if verdict.0 == "ok" {
            ok_all += 1;
            ok_real += real_doc as usize;
        } else {
            causes.entry(verdict.1.clone()).or_default().push(name.clone());
        }
        per_doc.insert(name, serde_json::json!({ "real": real_doc, "verdict": verdict.0 }));
    }
    let mut ranked: Vec<(&String, &Vec<String>)> = causes.iter().collect();
    ranked.sort_by_key(|(_, v)| std::cmp::Reverse(v.len()));
    println!(
        "[{}] rebuilt {}/{} (real {}/{})",
        if blank { "blank" } else { "template" },
        ok_all,
        files.len(),
        ok_real,
        real
    );
    for (cause, docs) in ranked.iter().take(40) {
        let real_n = docs.iter().filter(|d| is_real(d.rsplit('/').next().unwrap())).count();
        println!("  {:4} ({:3} real)  {}", docs.len(), real_n, cause);
    }
    if let Some(dir) = report.parent() {
        let _ = std::fs::create_dir_all(dir);
    }
    let summary = serde_json::json!({
        "mode": if blank { "blank" } else { "template" },
        "docs": files.len(), "rebuilt": ok_all, "real_docs": real, "real_rebuilt": ok_real,
        "causes": causes, "per_doc": per_doc,
    });
    std::fs::write(&report, serde_json::to_string_pretty(&summary).unwrap()).unwrap();
}
