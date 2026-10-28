//! oxi-lrpb-bench — check Oxi's page breaks against Word without running Word.
//!
//! Word writes `<w:lastRenderedPageBreak/>` (LRPB) wherever its own render of
//! the document started a new page, and `docProps/app.xml` `<Pages>` with the
//! page count. Both are Word's answer, saved in the file. This tool:
//!
//! 1. derives the page every paragraph and table row starts on from the LRPB
//!    markers (the ground truth), and checks the derived page count against
//!    `<Pages>` so a wrong derivation shows up instead of passing silently;
//! 2. lays the document out with `OXI_IGNORE_LRPB=1` ("blind": the parser drops
//!    the markers, so no layout rule can read the answer) and without it
//!    ("hinted": what Oxi does by default);
//! 3. scores both against the ground truth.
//!
//! The blind numbers are the honest ones. The hinted-minus-blind gap is how much
//! the saved hint is worth, and units the hint makes WRONG are layout bugs.
//!
//!     oxi-lrpb-bench <file.docx | dir>... [--json out.json] [--units]
//!
//! Only files that carry LRPB markers are scored (others have no ground truth).

use std::io::Read;
use std::path::{Path, PathBuf};

use oxidocs_core::ir::{Block, Document, Run, ShowRevisions, Table};
use oxidocs_core::layout::{LayoutContent, LayoutEngine, LayoutResult};
use serde_json::json;

const ENV: &str = "OXI_IGNORE_LRPB";

// ---------- ground truth from the markers ----------

/// One scored unit: a body paragraph, or one row of a top-level table.
struct Unit {
    container: usize,
    block: usize,
    row: Option<usize>,
    word_page: usize,
    text: String,
}

/// Markers in a run sequence: (before any visible text, after visible text).
fn marker_split<'a>(runs: impl Iterator<Item = &'a Run>, seen_text: &mut bool) -> (usize, usize) {
    let (mut before, mut after) = (0, 0);
    for r in runs {
        if r.has_last_rendered_page_break {
            // The marker sits ahead of the run's text in Word's output.
            if *seen_text { after += 1 } else { before += 1 }
        }
        if !r.text.trim().is_empty() {
            *seen_text = true;
        }
    }
    (before, after)
}

fn walk_blocks(blocks: &[Block], seen_text: &mut bool, acc: &mut (usize, usize)) {
    for b in blocks {
        match b {
            Block::Paragraph(p) => {
                let (s, m) = marker_split(p.runs.iter(), seen_text);
                acc.0 += s;
                acc.1 += m;
            }
            Block::Table(t) => {
                for row in &t.rows {
                    for c in &row.cells {
                        walk_blocks(&c.blocks, seen_text, acc);
                    }
                }
            }
            _ => {}
        }
    }
}

fn text_of(blocks: &[Block]) -> String {
    let mut s = String::new();
    for b in blocks {
        match b {
            Block::Paragraph(p) => p.runs.iter().for_each(|r| s.push_str(&r.text)),
            Block::Table(t) => t.rows.iter().flat_map(|r| r.cells.iter())
                .for_each(|c| s.push_str(&text_of(&c.blocks))),
            _ => {}
        }
    }
    s
}

/// Word marks a table row in two ways. A row Word SPLIT across pages carries a
/// marker after visible text in some cell (the continuation point). The first
/// row that STARTS on a new page carries one at the start of a cell, also when
/// the break came from splitting the row above; that one is a continuation
/// mark, not a second break.
fn table_units(t: &Table, container: usize, block: usize, page: &mut usize, out: &mut Vec<Unit>) {
    let mut prev_split = false;
    for (ri, row) in t.rows.iter().enumerate() {
        let (mut start, mut mid) = (false, false);
        for c in &row.cells {
            let mut seen = false;
            let mut acc = (0, 0);
            walk_blocks(&c.blocks, &mut seen, &mut acc);
            start |= acc.0 > 0;
            mid |= acc.1 > 0;
        }
        if start && !mid && !prev_split {
            *page += 1;
        }
        let text: String = row.cells.iter().map(|c| text_of(&c.blocks)).collect();
        out.push(Unit { container, block, row: Some(ri), word_page: *page, text });
        if mid {
            *page += 1;
        }
        prev_split = mid;
    }
}

fn ground_truth(doc: &Document) -> (Vec<Unit>, usize) {
    let mut page = 1;
    let mut out = vec![];
    for (ci, sec) in doc.pages.iter().enumerate() {
        for (bi, b) in sec.blocks.iter().enumerate() {
            match b {
                Block::Paragraph(p) => {
                    let mut seen = false;
                    let (before, after) = marker_split(p.runs.iter(), &mut seen);
                    page += before;
                    let text = p.runs.iter().map(|r| r.text.as_str()).collect();
                    out.push(Unit { container: ci, block: bi, row: None, word_page: page, text });
                    page += after;
                }
                Block::Table(t) => table_units(t, ci, bi, &mut page, &mut out),
                _ => {}
            }
        }
    }
    (out, page)
}

fn marker_total(doc: &Document) -> usize {
    let mut acc = (0, 0);
    for sec in &doc.pages {
        let mut seen = false;
        walk_blocks(&sec.blocks, &mut seen, &mut acc);
    }
    acc.0 + acc.1
}

// ---------- Oxi's pages ----------

fn norm(s: &str) -> String {
    s.chars().filter(|c| !c.is_whitespace()).collect()
}

/// First page holding a text fragment of this unit. Fragments are keyed by
/// (container, block, row) and must be part of the unit's own text, which keeps
/// header/footer fragments that share an index from matching.
fn oxi_page(res: &LayoutResult, u: &Unit) -> Option<usize> {
    let want = norm(&u.text);
    if want.is_empty() {
        return None;
    }
    for (pi, pg) in res.pages.iter().enumerate() {
        for el in &pg.elements {
            let LayoutContent::Text { text, .. } = &el.content else { continue };
            if el.paragraph_index != Some(u.block) {
                continue;
            }
            if el.source_container_index.is_some_and(|c| c != u.container) {
                continue;
            }
            if u.row.is_some() && el.cell_row_index != u.row {
                continue;
            }
            let t = norm(text);
            if !t.is_empty() && want.contains(&t) {
                return Some(pi + 1);
            }
        }
    }
    None
}

// ---------- driver ----------

fn app_pages(data: &[u8]) -> Option<usize> {
    let mut z = zip::ZipArchive::new(std::io::Cursor::new(data)).ok()?;
    let mut s = String::new();
    z.by_name("docProps/app.xml").ok()?.read_to_string(&mut s).ok()?;
    let i = s.find("<Pages>")? + 7;
    s[i..].split('<').next()?.trim().parse().ok()
}

/// The marker lives in the deflated `word/document.xml`, so look inside.
fn is_zip_with_marker(data: &[u8]) -> bool {
    let Ok(mut z) = zip::ZipArchive::new(std::io::Cursor::new(data)) else { return false };
    let mut s = String::new();
    let Ok(mut f) = z.by_name("word/document.xml") else { return false };
    f.read_to_string(&mut s).is_ok() && s.contains("lastRenderedPageBreak")
}

fn layout(data: &[u8], blind: bool) -> (Document, LayoutResult) {
    if blind {
        std::env::set_var(ENV, "1");
    } else {
        std::env::remove_var(ENV);
    }
    let doc = oxidocs_core::parse_docx(data).expect("parse");
    let res = LayoutEngine::for_document(&doc).with_show_revisions(ShowRevisions::Final).layout(&doc);
    std::env::remove_var(ENV);
    (doc, res)
}

fn collect(p: &Path, out: &mut Vec<PathBuf>) {
    if p.is_dir() {
        let mut entries: Vec<_> = std::fs::read_dir(p).into_iter().flatten().flatten().map(|e| e.path()).collect();
        entries.sort();
        for e in entries {
            collect(&e, out);
        }
    } else if p.extension().is_some_and(|e| e.eq_ignore_ascii_case("docx")) {
        out.push(p.to_path_buf());
    }
}

fn main() {
    let args: Vec<String> = std::env::args().skip(1).collect();
    let mut json_out = None;
    let mut show_units = false;
    let mut files = vec![];
    let mut it = args.iter();
    while let Some(a) = it.next() {
        match a.as_str() {
            "--json" => json_out = it.next().cloned(),
            "--units" => show_units = true,
            _ => collect(Path::new(a), &mut files),
        }
    }
    if files.is_empty() {
        eprintln!("usage: oxi-lrpb-bench <file.docx | dir>... [--json out.json] [--units]");
        std::process::exit(2);
    }

    let mut docs = vec![];
    let (mut skipped, mut failed) = (0, 0);
    let mut tot = [0usize; 6]; // units, blind_ok, hinted_ok, hint_hurts, word_breaks, blind_same_place
    let mut pages = [0usize; 4]; // oracle_ok, blind_count_ok, hinted_count_ok, docs
    let mut doc_pass = [0usize; 2];
    println!("{:<44} {:>4} {:>4} {:>5} {:>4} {:>6} {:>7} {:>7} {:>5}",
             "document", "app", "lrpb", "blind", "hint", "units", "blindOK", "hintOK", "hurts");
    for f in &files {
        let data = match std::fs::read(f) { Ok(d) => d, Err(_) => { failed += 1; continue } };
        // Cheap pre-check: no marker anywhere, no ground truth, no layout.
        if !data.windows(21).any(|w| w == b"lastRenderedPageBreak") && !is_zip_with_marker(&data) {
            skipped += 1;
            continue;
        }
        let run = std::panic::catch_unwind(|| {
            let (doc_h, res_h) = layout(&data, false);
            if marker_total(&doc_h) == 0 {
                return None;
            }
            let (doc_b, res_b) = layout(&data, true);
            assert_eq!(marker_total(&doc_b), 0, "{ENV} left markers in the IR");
            let (units, word_pages) = ground_truth(&doc_h);
            Some((units, word_pages, res_h, res_b))
        });
        let Ok(run) = run else { failed += 1; continue };
        let Some((units, word_pages, res_h, res_b)) = run else { skipped += 1; continue };

        let app = app_pages(&data);
        let mut rows = vec![];
        let (mut n, mut ok_b, mut ok_h, mut hurts) = (0, 0, 0, 0);
        let mut prev: Option<(usize, usize)> = None;
        let (mut breaks, mut same) = (0, 0);
        for u in &units {
            let (Some(b), Some(h)) = (oxi_page(&res_b, u), oxi_page(&res_h, u)) else { continue };
            n += 1;
            ok_b += (b == u.word_page) as usize;
            ok_h += (h == u.word_page) as usize;
            let hurt = b == u.word_page && h != u.word_page;
            hurts += hurt as usize;
            if let Some((pw, pb)) = prev {
                if u.word_page > pw {
                    breaks += 1;
                    same += (b > pb) as usize;
                }
            }
            prev = Some((u.word_page, b));
            if show_units || hurt {
                rows.push(json!({"container": u.container, "block": u.block, "row": u.row,
                    "word": u.word_page, "blind": b, "hinted": h, "hint_hurts": hurt,
                    "text": u.text.chars().take(40).collect::<String>()}));
            }
        }
        let name = f.file_name().unwrap().to_string_lossy().to_string();
        println!("{:<44} {:>4} {:>4} {:>5} {:>4} {:>6} {:>7} {:>7} {:>5}",
                 name.chars().take(44).collect::<String>(),
                 app.map(|a| a.to_string()).unwrap_or("-".into()), word_pages,
                 res_b.pages.len(), res_h.pages.len(), n, ok_b, ok_h, hurts);
        tot[0] += n; tot[1] += ok_b; tot[2] += ok_h; tot[3] += hurts; tot[4] += breaks; tot[5] += same;
        pages[3] += 1;
        pages[0] += (app == Some(word_pages)) as usize;
        pages[1] += (res_b.pages.len() == word_pages) as usize;
        pages[2] += (res_h.pages.len() == word_pages) as usize;
        doc_pass[0] += (n > 0 && ok_b == n) as usize;
        doc_pass[1] += (n > 0 && ok_h == n) as usize;
        docs.push(json!({"file": f.to_string_lossy(), "app_pages": app, "word_pages_from_lrpb": word_pages,
            "oxi_pages_blind": res_b.pages.len(), "oxi_pages_hinted": res_h.pages.len(),
            "units": n, "blind_ok": ok_b, "hinted_ok": ok_h, "hint_hurts": hurts,
            "word_breaks": breaks, "blind_same_place": same, "detail": rows}));
    }

    let summary = json!({
        "docs_scored": pages[3], "docs_without_markers": skipped, "docs_failed": failed,
        "oracle_matches_app_xml_pages": pages[0],
        "units": tot[0], "blind_ok": tot[1], "hinted_ok": tot[2], "hint_hurts": tot[3],
        "page_count_match_blind": pages[1], "page_count_match_hinted": pages[2],
        "docs_all_units_ok_blind": doc_pass[0], "docs_all_units_ok_hinted": doc_pass[1],
        "word_breaks": tot[4], "blind_breaks_same_place": tot[5],
    });
    println!("\n{}", serde_json::to_string_pretty(&summary).unwrap());
    if let Some(p) = json_out {
        std::fs::write(p, serde_json::to_string_pretty(&json!({"summary": summary, "docs": docs})).unwrap()).unwrap();
    }
}
