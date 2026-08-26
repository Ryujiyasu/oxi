// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! Does opening a document, changing one thing, and saving it leave the rest
//! of the document alone?
//!
//! Every metric this project keeps asks how well a file is DRAWN. None of
//! them asks whether it survives being written back, and an editor that
//! quietly drops a sheet's charts or a document's numbering is worse than one
//! that draws them a pixel out. The methodology has named round-trip
//! preservation as a metric since the phases were laid out; this is it.
//!
//! The test is an edit that asks for nothing to change: the first cell that
//! holds something is written back with the value and type it already has.
//! The
//! editor still takes the full save path — it patches the archive, rebuilds
//! the parts it touches and hands back new bytes — so anything the writer
//! loses, it loses here. What comes back must therefore parse to exactly the
//! document that went in.
//!
//! Reported per file, and as three counts over the corpus: how many opened,
//! how many saved, and how many came back whole.
//!
//! A document's editor counts paragraphs as the body's own top-level
//! children, which is not how the IR lays them out, so the run to write back
//! is chosen through `DocxEditor::addressable_runs` — the editor's own view of
//! what it can address — rather than guessed at from the IR.
//!
//!     oxi-roundtrip <file-or-directory> [--quiet] [--limit N]

use std::path::{Path, PathBuf};

use oxicells_core::{parse_xlsx, CellEditValue, XlsxEditor};
use oxidocs_core::{parse_docx, DocxEditor};
use oxislides_core::{parse_pptx, PptxEditor};

/// What became of one file.
enum Verdict {
    /// Nothing in it could be edited, so it says nothing about the writer.
    Untested(&'static str),
    Whole,
    /// It opened, saved and reopened, but came back different.
    Changed(String),
    /// It would not open, or would not save.
    Broken(String),
}

fn main() {
    let mut args = std::env::args().skip(1);
    let mut where_from: Option<PathBuf> = None;
    let mut quiet = false;
    let mut limit = usize::MAX;
    // Where to keep what was written, so that Office can be asked whether it
    // accepts it. An IR that matches proves the writer kept everything WE
    // model; it cannot prove the file still opens.
    let mut keep: Option<PathBuf> = None;
    // Write a mark instead of the value that is already there. An edit that
    // changes nothing proves the writer breaks nothing; it does not prove an
    // edit ARRIVES. With this the tool says where it aimed, and Office can be
    // asked whether that is the one cell that moved.
    let mut sentinel: Option<String> = None;
    while let Some(arg) = args.next() {
        match arg.as_str() {
            "--quiet" => quiet = true,
            "--limit" => limit = args.next().and_then(|n| n.parse().ok()).unwrap_or(usize::MAX),
            "--keep" => keep = args.next().map(PathBuf::from),
            "--sentinel" => sentinel = Some(args.next().unwrap_or_else(|| "OXI".into())),
            _ => where_from = Some(PathBuf::from(arg)),
        }
    }
    let Some(where_from) = where_from else {
        eprintln!(
                "usage: oxi-roundtrip <file-or-directory> [--quiet] [--limit N] [--keep DIR] [--sentinel TEXT]"
            );
        std::process::exit(2);
    };

    let mut files: Vec<PathBuf> = Vec::new();
    if where_from.is_dir() {
        let mut held: Vec<PathBuf> = std::fs::read_dir(&where_from)
            .unwrap_or_else(|trouble| panic!("cannot read {}: {trouble}", where_from.display()))
            .filter_map(|entry| entry.ok().map(|entry| entry.path()))
            .filter(|path| matches!(kind(path), Some(_)))
            .collect();
        held.sort();
        files.extend(held.into_iter().take(limit));
    } else {
        files.push(where_from);
    }

    let (mut opened, mut saved, mut whole, mut untested) = (0, 0, 0, 0);
    let mut wrong: Vec<(PathBuf, String)> = Vec::new();
    for path in &files {
        let verdict = match kind(path) {
            Some(Kind::Xlsx) => xlsx(path, keep.as_deref(), sentinel.as_deref()),
            Some(Kind::Docx) => docx(path, keep.as_deref(), sentinel.as_deref()),
            Some(Kind::Pptx) => pptx(path, keep.as_deref(), sentinel.as_deref()),
            None => continue,
        };
        match &verdict {
            Verdict::Untested(why) => {
                untested += 1;
                opened += 1;
                if !quiet {
                    println!("  --  {}  ({why})", name(path));
                }
            }
            Verdict::Whole => {
                opened += 1;
                saved += 1;
                whole += 1;
                if !quiet {
                    println!("  ok  {}", name(path));
                }
            }
            Verdict::Changed(what) => {
                opened += 1;
                saved += 1;
                wrong.push((path.clone(), what.clone()));
                println!("  !!  {}  {what}", name(path));
            }
            Verdict::Broken(what) => {
                wrong.push((path.clone(), what.clone()));
                println!("  XX  {}  {what}", name(path));
            }
        }
    }

    let tested = files.len() - untested;
    println!();
    println!("  {} file(s): {opened} opened, {saved} saved, {whole} came back whole", files.len());
    if untested > 0 {
        println!("  {untested} held nothing to edit and are not counted in the {tested} tested");
    }
    if !wrong.is_empty() {
        println!("  {} did not survive:", wrong.len());
        for (path, what) in wrong.iter().take(20) {
            println!("    {}  {what}", name(path));
        }
    }
}

enum Kind {
    Xlsx,
    Docx,
    Pptx,
}

fn kind(path: &Path) -> Option<Kind> {
    // A file Office has left open is mirrored as `~$name`; it is a lock, not a
    // document, and it parses as rubbish.
    if name(path).starts_with("~$") {
        return None;
    }
    match path.extension().and_then(|held| held.to_str()) {
        Some("xlsx") | Some("xlsm") => Some(Kind::Xlsx),
        Some("docx") => Some(Kind::Docx),
        Some("pptx") => Some(Kind::Pptx),
        _ => None,
    }
}

fn name(path: &Path) -> String {
    path.file_name()
        .map(|held| held.to_string_lossy().to_string())
        .unwrap_or_default()
}

/// Where the two documents part company, in a form that names the place.
///
/// Comparing the serialised trees whole says only "different", which is no
/// use across two hundred workbooks. This walks them together and reports the
/// first path that differs, so a run of failures can be read as one cause.
fn parted(before: &serde_json::Value, after: &serde_json::Value, at: &str) -> Option<String> {
    use serde_json::Value;
    match (before, after) {
        (Value::Object(one), Value::Object(two)) => {
            for (key, held) in one {
                match two.get(key) {
                    None => return Some(format!("{at}.{key} is gone")),
                    Some(other) => {
                        if let Some(found) = parted(held, other, &format!("{at}.{key}")) {
                            return Some(found);
                        }
                    }
                }
            }
            two.keys()
                .find(|key| !one.contains_key(*key))
                .map(|key| format!("{at}.{key} appeared"))
        }
        (Value::Array(one), Value::Array(two)) => {
            if one.len() != two.len() {
                return Some(format!("{at} held {} and now holds {}", one.len(), two.len()));
            }
            one.iter()
                .zip(two)
                .enumerate()
                .find_map(|(step, (held, other))| parted(held, other, &format!("{at}[{step}]")))
        }
        _ => {
            if before == after {
                None
            } else {
                Some(format!("{at}: {} became {}", brief(before), brief(after)))
            }
        }
    }
}

/// Every place the two differ, up to a handful.
///
/// `parted` stops at the first, which is what a same-or-not question wants. A
/// MARK is a difference on purpose, so the question becomes how many there
/// are: exactly one, holding the mark, means the edit went where it was aimed
/// and nowhere else.
fn differences(
    before: &serde_json::Value,
    after: &serde_json::Value,
    at: &str,
    found: &mut Vec<(String, String)>,
) {
    use serde_json::Value;
    if found.len() > 8 {
        return;
    }
    match (before, after) {
        (Value::Object(one), Value::Object(two)) => {
            for (key, held) in one {
                match two.get(key) {
                    None => found.push((format!("{at}.{key}"), "is gone".into())),
                    Some(other) => differences(held, other, &format!("{at}.{key}"), found),
                }
            }
            for (key, held) in two.iter().filter(|(key, _)| !one.contains_key(*key)) {
                // What appeared, not the fact that something did: a value
                // written over one of another type arrives as a NEW key —
                // `{"Number":1}` becomes `{"String":"OXIMARK"}` — so saying
                // only "appeared" loses the very thing being looked for.
                found.push((format!("{at}.{key}"), brief(held)));
            }
        }
        (Value::Array(one), Value::Array(two)) => {
            if one.len() != two.len() {
                found.push((at.to_string(), format!("{} became {}", one.len(), two.len())));
                return;
            }
            for (step, (held, other)) in one.iter().zip(two).enumerate() {
                differences(held, other, &format!("{at}[{step}]"), found);
            }
        }
        _ => {
            if before != after {
                found.push((at.to_string(), brief(after)));
            }
        }
    }
}

/// The element a path belongs to: everything up to and including its last
/// index.
///
/// Counting differing LEAVES asks the wrong question. Writing a string over a
/// number turns `{"Number":1}` into `{"String":"OXIMARK"}`, which is one key
/// gone and one appeared — two leaves, one cell. Writing a value over a
/// formula changes the value and takes the formula away — two leaves, one
/// cell. Neither is a leak. What a leak looks like is two DIFFERENT cells
/// moving, and that is what this makes visible.
fn element_of(path: &str) -> &str {
    match path.rfind(']') {
        Some(at) => &path[..=at],
        None => path,
    }
}

/// The verdict for a marked edit: one element changed, and it holds the mark.
fn only_the_mark<T: serde::Serialize>(before: &T, after: &T, mark: &str) -> Verdict {
    let (Ok(one), Ok(two)) = (
        serde_json::to_value(before),
        serde_json::to_value(after),
    ) else {
        return Verdict::Broken("cannot be compared".into());
    };
    let mut found = Vec::new();
    differences(&one, &two, "", &mut found);
    if found.is_empty() {
        return Verdict::Changed("the mark never arrived".into());
    }
    let mut elements: Vec<&str> = found.iter().map(|(at, _)| element_of(at)).collect();
    elements.sort_unstable();
    elements.dedup();
    if elements.len() > 1 {
        return Verdict::Changed(format!(
            "{} elements changed, not one: {}",
            elements.len(),
            elements.join(", ")
        ));
    }
    if !found.iter().any(|(_, held)| held.contains(mark)) {
        return Verdict::Changed(format!(
            "one element changed, but nothing in it holds the mark: {}",
            found[0].0
        ));
    }
    Verdict::Whole
}

fn brief(held: &serde_json::Value) -> String {
    let shown = held.to_string();
    if shown.chars().count() > 40 {
        format!("{}…", shown.chars().take(40).collect::<String>())
    } else {
        shown
    }
}

fn xlsx(path: &Path, keep: Option<&Path>, sentinel: Option<&str>) -> Verdict {
    let Ok(bytes) = std::fs::read(path) else {
        return Verdict::Broken("cannot be read".into());
    };
    let before = match parse_xlsx(&bytes) {
        Ok(held) => held,
        Err(trouble) => return Verdict::Broken(format!("will not open: {trouble}")),
    };
    // The first cell whose own content the editor can be ASKED for.
    //
    // Not every cell can: a cell dressed in several fonts holds its stretches
    // in `runs`, and there is no edit that says "this text, dressed as it
    // was" — `CellEditValue::String` flattens it. Writing one back is a real
    // change, so counting it as a lost round trip would be blaming the writer
    // for doing as it was told. A cell carrying a formula CAN be asked for:
    // it is written back as the formula, not as the value Excel cached for it,
    // which is the stronger test of the two.
    let mut asked = None;
    let mut dressed = 0usize;
    'hunt: for (sheet_at, sheet) in before.sheets.iter().enumerate() {
        for row in &sheet.rows {
            for cell in &row.cells {
                if !cell.runs.is_empty() {
                    dressed += 1;
                    continue;
                }
                let value = match (&cell.formula, &cell.value) {
                    (Some(held), _) => CellEditValue::Formula(held.clone()),
                    (None, oxicells_core::ir::CellValue::String(held)) => {
                        CellEditValue::String(held.clone())
                    }
                    (None, oxicells_core::ir::CellValue::Number(held)) => {
                        CellEditValue::Number(*held)
                    }
                    (None, oxicells_core::ir::CellValue::Boolean(held)) => {
                        CellEditValue::Boolean(*held)
                    }
                    _ => continue,
                };
                asked = Some((sheet_at, row.index, cell.col, value));
                break 'hunt;
            }
        }
    }
    let Some((sheet_at, row, col, value)) = asked else {
        return Verdict::Untested(if dressed > 0 {
            "every cell it holds is dressed"
        } else {
            "no cell holds a value"
        });
    };
    let mut editor = match XlsxEditor::new(&bytes) {
        Ok(held) => held,
        Err(trouble) => return Verdict::Broken(format!("will not open to edit: {trouble}")),
    };
    match sentinel {
        None => editor.set_cell_value(sheet_at, row, col, value),
        Some(mark) => {
            editor.set_cell_value(sheet_at, row, col, CellEditValue::String(mark.to_string()));
            println!("  aimed {} sheet {sheet_at} row {row} col {col}", name(path));
        }
    }
    let written = match editor.save() {
        Ok(held) => held,
        Err(trouble) => return Verdict::Broken(format!("will not save: {trouble}")),
    };
    kept(keep, path, &written);
    let after = match parse_xlsx(&written) {
        Ok(held) => held,
        Err(trouble) => return Verdict::Changed(format!("will not reopen: {trouble}")),
    };
    if let Some(mark) = sentinel {
        return only_the_mark(&before, &after, mark);
    }
    // Sheet by sheet, not workbook by workbook. Serialising both trees whole
    // took five gigabytes on the corpus's largest books — a metric nobody can
    // afford to run is one nobody runs.
    if before.sheets.len() != after.sheets.len() {
        return Verdict::Changed(format!(
            "held {} sheet(s) and now holds {}",
            before.sheets.len(),
            after.sheets.len()
        ));
    }
    for (at, (one, two)) in before.sheets.iter().zip(&after.sheets).enumerate() {
        if let Verdict::Changed(what) = compare_at(one, two, &format!(".sheets[{at}]")) {
            return Verdict::Changed(what);
        }
    }
    compare_at(&before.default_style, &after.default_style, ".default_style")
}

fn docx(path: &Path, keep: Option<&Path>, sentinel: Option<&str>) -> Verdict {
    let Ok(bytes) = std::fs::read(path) else {
        return Verdict::Broken("cannot be read".into());
    };
    let before = match parse_docx(&bytes) {
        Ok(held) => held,
        Err(trouble) => return Verdict::Broken(format!("will not open: {trouble}")),
    };
    let mut editor = match DocxEditor::new(&bytes) {
        Ok(held) => held,
        Err(trouble) => return Verdict::Broken(format!("will not open to edit: {trouble}")),
    };
    // The first run carrying text, counted the way the editor counts.
    let addressable = match editor.addressable_runs() {
        Ok(held) => held,
        Err(trouble) => return Verdict::Broken(format!("will not say what it holds: {trouble}")),
    };
    let mut asked = None;
    'hunt: for (para_at, runs) in addressable.iter().enumerate() {
        for (run_at, text) in runs.iter().enumerate() {
            if !text.is_empty() {
                asked = Some((para_at, run_at, text.clone()));
                break 'hunt;
            }
        }
    }
    match asked {
        Some((para_at, run_at, text)) => match sentinel {
            None => editor.set_run_text(para_at, run_at, text),
            Some(mark) => {
                editor.set_run_text(para_at, run_at, mark.to_string());
                println!("  aimed {} body paragraph {para_at} run {run_at}", name(path));
            }
        },
        None => {
            // Fifteen percent of the corpus keeps all its text inside tables,
            // which no body run reaches. Those documents are edited through
            // the cells, and leaving them untested left the writer's whole
            // table path unmeasured.
            let cells = match editor.addressable_cells() {
                Ok(held) => held,
                Err(trouble) => {
                    return Verdict::Broken(format!("will not say what it holds: {trouble}"))
                }
            };
            // One RUN inside a cell, not the whole cell. `set_cell_text`
            // takes a single string for the cell, so handing a cell of several
            // runs its own text back still flattens it — which left every
            // document whose cells are dressed untestable, and they are the
            // ones most worth testing.
            let mut found = None;
            'hunt: for (table_at, rows) in cells.iter().enumerate() {
                for (row_at, row) in rows.iter().enumerate() {
                    for (col_at, cell) in row.iter().enumerate() {
                        for (para_at, runs) in cell.paragraphs.iter().enumerate() {
                            for (run_at, text) in runs.iter().enumerate() {
                                if !text.is_empty() {
                                    found = Some((
                                        table_at,
                                        row_at,
                                        col_at,
                                        para_at,
                                        run_at,
                                        text.clone(),
                                    ));
                                    break 'hunt;
                                }
                            }
                        }
                    }
                }
            }
            let Some((table_at, row_at, col_at, para_at, run_at, text)) = found else {
                return Verdict::Untested("nothing in it carries text");
            };
            match sentinel {
                None => {
                    editor.set_table_run_text(table_at, row_at, col_at, para_at, run_at, &text)
                }
                Some(mark) => {
                    editor.set_table_run_text(table_at, row_at, col_at, para_at, run_at, mark);
                    println!(
                        "  aimed {} table {table_at} row {row_at} col {col_at} paragraph {para_at} run {run_at}",
                        name(path)
                    );
                }
            }
        }
    }
    let written = match editor.save() {
        Ok(held) => held,
        Err(trouble) => return Verdict::Broken(format!("will not save: {trouble}")),
    };
    kept(keep, path, &written);
    let after = match parse_docx(&written) {
        Ok(held) => held,
        Err(trouble) => return Verdict::Changed(format!("will not reopen: {trouble}")),
    };
    if let Some(mark) = sentinel {
        // A mark is a difference on purpose, so the question is how many.
        // Exactly one, holding the mark, is an edit that went where it was
        // aimed and took nothing with it. Asking our own reader rather than
        // Word costs the outside opinion but gains an answer for every
        // document: Word hangs without warning on the ones that link out.
        return only_the_mark(&before, &after, mark);
    }
    compare(&before, &after)
}

/// Put what was written where Office can be pointed at it.
fn kept(keep: Option<&Path>, path: &Path, written: &[u8]) {
    let Some(keep) = keep else { return };
    let _ = std::fs::create_dir_all(keep);
    let _ = std::fs::write(keep.join(name(path)), written);
}

fn pptx(path: &Path, keep: Option<&Path>, sentinel: Option<&str>) -> Verdict {
    let Ok(bytes) = std::fs::read(path) else {
        return Verdict::Broken("cannot be read".into());
    };
    let before = match parse_pptx(&bytes) {
        Ok(held) => held,
        Err(trouble) => return Verdict::Broken(format!("will not open: {trouble}")),
    };
    let mut editor = match PptxEditor::new(&bytes) {
        Ok(held) => held,
        Err(trouble) => return Verdict::Broken(format!("will not open to edit: {trouble}")),
    };
    let addressable = match editor.addressable_runs() {
        Ok(held) => held,
        Err(trouble) => return Verdict::Broken(format!("will not say what it holds: {trouble}")),
    };
    let mut asked = None;
    'hunt: for (slide_at, shapes) in addressable.iter().enumerate() {
        for (shape_at, paragraphs) in shapes.iter().enumerate() {
            for (para_at, runs) in paragraphs.iter().enumerate() {
                for (run_at, text) in runs.iter().enumerate() {
                    if !text.is_empty() {
                        asked = Some((slide_at, shape_at, para_at, run_at, text.clone()));
                        break 'hunt;
                    }
                }
            }
        }
    }
    let Some((slide_at, shape_at, para_at, run_at, text)) = asked else {
        return Verdict::Untested("no run carries text");
    };
    match sentinel {
        None => editor.set_run_text(slide_at, shape_at, para_at, run_at, text),
        Some(mark) => {
            editor.set_run_text(slide_at, shape_at, para_at, run_at, mark.to_string());
            println!(
                "  aimed {} slide {slide_at} shape {shape_at} paragraph {para_at} run {run_at}",
                name(path)
            );
        }
    }
    let written = match editor.save() {
        Ok(held) => held,
        Err(trouble) => return Verdict::Broken(format!("will not save: {trouble}")),
    };
    kept(keep, path, &written);
    let after = match parse_pptx(&written) {
        Ok(held) => held,
        Err(trouble) => return Verdict::Changed(format!("will not reopen: {trouble}")),
    };
    if let Some(mark) = sentinel {
        return only_the_mark(&before, &after, mark);
    }
    if before.slides.len() != after.slides.len() {
        return Verdict::Changed(format!(
            "held {} slide(s) and now holds {}",
            before.slides.len(),
            after.slides.len()
        ));
    }
    for (at, (one, two)) in before.slides.iter().zip(&after.slides).enumerate() {
        if let Verdict::Changed(what) = compare_at(one, two, &format!(".slides[{at}]")) {
            return Verdict::Changed(what);
        }
    }
    Verdict::Whole
}

fn compare<T: serde::Serialize>(before: &T, after: &T) -> Verdict {
    compare_at(before, after, "")
}

fn compare_at<T: serde::Serialize>(before: &T, after: &T, at: &str) -> Verdict {
    let (Ok(one), Ok(two)) = (
        serde_json::to_value(before),
        serde_json::to_value(after),
    ) else {
        return Verdict::Broken("cannot be compared".into());
    };
    match parted(&one, &two, at) {
        None => Verdict::Whole,
        Some(what) => Verdict::Changed(what),
    }
}
