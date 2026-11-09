//! How often does the renderer's paragraph text (oxidocs-core) equal the
//! text the model says a paragraph shows (oxidocs-model `shown_positions`)?
//! The editor maps a click on the page to a model position through that
//! text, so this measures how much of it maps directly.
//!
//!     cargo run -p oxidocs-wasm --release --example map_census -- <docx or dir>... [--show N]
//!
//! Per document: the core paragraphs in document order (through table
//! cells), the model paragraphs in its numbering, and how many core
//! paragraphs find a model paragraph with the same text in order (a
//! longest common subsequence). `--show N` prints the first N core
//! paragraphs that find none, next to the model paragraph at that place.

use std::path::{Path, PathBuf};

use oxidocs_core::ir::{Block, Document};
use oxidocs_model::edit::{paragraph_texts, shown_positions, Story};
use oxidocs_model::DocxModel;

fn docs(p: &Path, out: &mut Vec<PathBuf>) {
    if p.is_dir() {
        let mut v: Vec<_> = std::fs::read_dir(p).into_iter().flatten().flatten().map(|e| e.path()).collect();
        v.sort();
        for e in v {
            docs(&e, out);
        }
    } else if p.extension().is_some_and(|x| x == "docx") && !p.file_name().unwrap().to_string_lossy().starts_with("~$") {
        out.push(p.to_path_buf());
    }
}

fn core_paragraphs(doc: &Document) -> Vec<String> {
    fn walk(blocks: &[Block], out: &mut Vec<String>) {
        for b in blocks {
            match b {
                Block::Paragraph(p) => out.push(p.runs.iter().map(|r| r.text.as_str()).collect()),
                Block::Table(t) => {
                    for row in &t.rows {
                        for cell in &row.cells {
                            walk(&cell.blocks, out);
                        }
                    }
                }
                _ => {}
            }
        }
    }
    let mut out = Vec::new();
    for page in &doc.pages {
        walk(&page.blocks, &mut out);
    }
    out
}

fn lcs(a: &[String], b: &[String]) -> (usize, Vec<bool>) {
    let (n, m) = (a.len(), b.len());
    let mut t = vec![vec![0u32; m + 1]; n + 1];
    for i in (0..n).rev() {
        for j in (0..m).rev() {
            t[i][j] = if a[i] == b[j] { t[i + 1][j + 1] + 1 } else { t[i + 1][j].max(t[i][j + 1]) };
        }
    }
    let mut matched = vec![false; n];
    let (mut i, mut j) = (0, 0);
    while i < n && j < m {
        if a[i] == b[j] {
            matched[i] = true;
            i += 1;
            j += 1;
        } else if t[i + 1][j] >= t[i][j + 1] {
            i += 1;
        } else {
            j += 1;
        }
    }
    (t[0][0] as usize, matched)
}

fn main() {
    let mut args: Vec<String> = std::env::args().skip(1).collect();
    let show = args.iter().position(|a| a == "--show").map(|i| {
        let n = args[i + 1].parse().unwrap();
        args.drain(i..i + 2);
        n
    }).unwrap_or(0);
    let mut files = Vec::new();
    for a in &args {
        docs(Path::new(a), &mut files);
    }
    let (mut total_core, mut total_matched, mut full_docs, mut n_docs) = (0, 0, 0, 0);
    for f in files {
        let Ok(bytes) = std::fs::read(&f) else { continue };
        let Ok(core) = oxidocs_core::parse_docx(&bytes) else { continue };
        let Ok(model) = DocxModel::read(&bytes) else { continue };
        let Ok(texts) = paragraph_texts(&model, &Story::Body) else { continue };
        let shown: Vec<String> = texts.iter().map(|t| shown_positions(t).into_iter().map(|(c, _)| c).collect()).collect();
        let core_texts = core_paragraphs(&core);
        if core_texts.len() > 4000 || shown.len() > 4000 {
            continue; // the quadratic table would not fit; rare
        }
        let (k, matched) = lcs(&core_texts, &shown);
        n_docs += 1;
        total_core += core_texts.len();
        total_matched += k;
        if k == core_texts.len() {
            full_docs += 1;
        }
        println!("{:5}/{:5} core paragraphs match (model {:5})  {}", k, core_texts.len(), shown.len(), f.display());
        let mut left = show;
        for (i, ok) in matched.iter().enumerate() {
            if left == 0 {
                break;
            }
            if !ok {
                left -= 1;
                let near = shown.get(i).map(|s| s.as_str()).unwrap_or("-");
                println!("      core[{i}] {:?}\n      model   {:?}", core_texts[i].chars().take(60).collect::<String>(), near.chars().take(60).collect::<String>());
            }
        }
    }
    println!("docs {n_docs}, all core paragraphs matched in {full_docs}; paragraphs {total_matched}/{total_core}");
}
