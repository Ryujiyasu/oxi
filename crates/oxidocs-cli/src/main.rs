// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

// PdfBuilder scaffold + render_* helpers below are retained as a reference
// implementation; the active path is docx_to_pdf → doc_to_pdf → oxipdf_core.
#![allow(dead_code, unused_variables, unused_imports)]

use std::collections::HashMap;
use std::fs;
use oxidocs_core::ir::{Block, Paragraph, Table, Alignment, TextBox, FloatingPosition};
use oxidocs_core::font::FontMetricsRegistry;
use oxipdf_core::ir::*;

mod fontidx;
mod vba;
mod emf;
mod raster;

#[cfg(target_env = "musl")]
#[global_allocator]
static ALLOCATOR: mimalloc::MiMalloc = mimalloc::MiMalloc;

/// Set by `--quiet`: progress notes and the closing "Converted" line are
/// left out. Warnings and errors are still written.
static QUIET: std::sync::atomic::AtomicBool = std::sync::atomic::AtomicBool::new(false);

fn quiet() -> bool {
    QUIET.load(std::sync::atomic::Ordering::Relaxed)
}

/// A progress note on stderr: what was parsed, which fonts were embedded.
macro_rules! note {
    ($($arg:tt)*) => {
        if !quiet() {
            eprintln!($($arg)*);
        }
    };
}

fn main() {
    let mut args: Vec<String> = std::env::args().collect();
    if let Some(at) = args.iter().skip(2).position(|a| a == "--quiet" || a == "-q") {
        args.remove(at + 2);
        QUIET.store(true, std::sync::atomic::Ordering::Relaxed);
    }

    match args.get(1).map(|s| s.as_str()) {
        Some("--version") | Some("-V") | Some("version") => {
            println!("oxidocs {}", env!("CARGO_PKG_VERSION"));
        }
        Some("docx-to-pdf") => {
            let usage = "Usage: oxi docx-to-pdf [--strict] <input.docx> <output.pdf>";
            let strict = args[2..].iter().any(|a| a == "--strict");
            let paths: Vec<&String> = args[2..].iter().filter(|a| *a != "--strict").collect();
            let input = paths.first().expect(usage);
            let output = paths.get(1).expect(usage);
            docx_to_pdf(input, output, strict);
        }
        Some("docx-to-png") => {
            let usage = "Usage: oxi docx-to-png [--strict] [--page N] [--width PX] [--height PX] <input.docx> <output.png>";
            let mut strict = false;
            let mut page = 1usize;
            let mut size = raster::Size { width: None, height: None };
            let mut paths = Vec::new();
            let mut rest = args[2..].iter();
            while let Some(arg) = rest.next() {
                let mut number = || rest.next().and_then(|v| v.parse::<u32>().ok()).filter(|n| *n > 0).expect(usage);
                match arg.as_str() {
                    "--strict" => strict = true,
                    "--page" => page = number() as usize,
                    "--width" => size.width = Some(number()),
                    "--height" => size.height = Some(number()),
                    _ => paths.push(arg.as_str()),
                }
            }
            let (Some(input), Some(output)) = (paths.first(), paths.get(1)) else {
                eprintln!("{usage}");
                std::process::exit(1);
            };
            finish(input, output, convert_png(input, output, strict, page, size));
        }
        Some("serve") => {
            let threads = args[2..]
                .windows(2)
                .find(|w| w[0] == "--threads")
                .and_then(|w| w[1].parse::<usize>().ok())
                .unwrap_or(1)
                .max(1);
            serve(threads);
        }
        Some("fonts-index") => {
            // Instrument for the font resolver: report what the machine
            // actually offers, so a substitution decision can be read
            // rather than guessed.
            let index = fontidx::FontIndex::build(&bundled_font_dirs());
            println!("faces indexed: {}", index.len());
            let query = args.get(2).map(|s| s.as_str());
            let families: Vec<&str> = match query {
                Some(q) => vec![q],
                None => vec![
                    "Calibri", "Cambria", "Arial", "Times New Roman", "Courier New",
                    "Verdana", "Tahoma", "Segoe UI", "Georgia", "Wingdings", "Symbol",
                    "Carlito", "Liberation Sans", "Liberation Serif", "Liberation Mono",
                ],
            };
            for family in families {
                for (bold, italic) in [(false, false), (true, false), (false, true), (true, true)] {
                    let style = match (bold, italic) {
                        (false, false) => "regular",
                        (true, false) => "bold",
                        (false, true) => "italic",
                        (true, true) => "bold-italic",
                    };
                    match index.find(family, bold, italic) {
                        Some(face) if face.bold == bold && face.italic == italic => println!(
                            "  {family:<18} {style:<12} {}#{}",
                            face.path.display(),
                            face.index
                        ),
                        Some(face) => println!(
                            "  {family:<18} {style:<12} DEGRADED -> {} ({}{})",
                            face.path.display(),
                            if face.bold { "bold" } else { "regular" },
                            if face.italic { "+italic" } else { "" }
                        ),
                        None => println!("  {family:<18} {style:<12} ABSENT"),
                    }
                }
            }
        }
        Some("vba-analyze") => {
            let input = args.get(2).expect("Usage: oxi vba-analyze <input.xlsm>");
            if let Err(error) = vba::analyze_file(input) {
                eprintln!("VBA analysis failed: {error}");
                std::process::exit(2);
            }
        }
        Some("vba-safety") => {
            let input = args.get(2).expect("Usage: oxi vba-safety <input.xlsm>");
            match vba::safety(input) {
                Ok(true) => std::process::exit(3),
                Ok(false) => {}
                Err(error) => {
                    eprintln!("VBA safety read failed: {error}");
                    std::process::exit(2);
                }
            }
        }
        Some("hanko-keygen") => {
            let path = args
                .get(2)
                .expect("Usage: oxi hanko-keygen <key.pk8>");
            if let Err(error) = vba::keygen(path) {
                eprintln!("Key generation failed: {error}");
                std::process::exit(2);
            }
        }
        Some("vba-seal") => {
            let input = args
                .get(2)
                .expect("Usage: oxi vba-seal <input.xlsm> <key.pk8> [note]");
            let key = args
                .get(3)
                .expect("Usage: oxi vba-seal <input.xlsm> <key.pk8> [note]");
            let note = args.get(4).map(String::as_str).unwrap_or("");
            if let Err(error) = vba::seal(input, key, note) {
                eprintln!("VBA sealing failed: {error}");
                std::process::exit(2);
            }
        }
        Some("vba-inventory") => {
            let input = args
                .get(2)
                .expect("Usage: oxi vba-inventory <file-or-directory>");
            if let Err(error) = vba::inventory(input) {
                eprintln!("VBA inventory failed: {error}");
                std::process::exit(2);
            }
        }
        Some("vba-inventory-json") => {
            let input = args
                .get(2)
                .expect("Usage: oxi vba-inventory-json <file-or-directory>");
            if let Err(error) = vba::inventory_json(input) {
                eprintln!("VBA JSON inventory failed: {error}");
                std::process::exit(2);
            }
        }
        _ => {
            eprintln!("Oxi CLI - Document processing toolkit");
            eprintln!();
            eprintln!("Usage:");
            eprintln!("  oxi docx-to-pdf [--strict] [--quiet] <input.docx> <output.pdf>");
            eprintln!("  oxi docx-to-png [--strict] [--quiet] [--page N] [--width PX] [--height PX] <input.docx> <output.png>");
            eprintln!("      page N (default 1), fitted inside the given width and/or height (default 96 dpi)");
            eprintln!("  oxi serve [--threads N]");
            eprintln!("      one JSON request per stdin line, {{\"input\", \"output\", \"strict\", \"id\"}};");
            eprintln!("      an output ending in .png draws one page: add \"page\", \"width\", \"height\"");
            eprintln!("      one JSON answer per stdout line, in completion order");
            eprintln!("      --strict: write nothing and exit {EXIT_UNSUPPORTED} when the document holds");
            eprintln!("                content Oxi does not draw (listed as JSON on stderr)");
            eprintln!("      --quiet:  print nothing on success: no progress notes, no closing line");
            eprintln!("                (also for serve); warnings and errors are still written");
            eprintln!("  oxi --version");
            eprintln!("  oxi vba-analyze <input.xlsm>");
            eprintln!("  oxi vba-safety <input.xlsm>");
            eprintln!("  oxi vba-seal <input.xlsm> <key.pk8> [note]");
            eprintln!("  oxi hanko-keygen <key.pk8>");
            eprintln!("  oxi vba-inventory <file-or-directory>");
            eprintln!("  oxi vba-inventory-json <file-or-directory>");
            std::process::exit(1);
        }
    }
}

/// `docx-to-pdf --strict` exit status when the document holds content Oxi does
/// not draw. Distinct from usage errors (1) and panics (101), so a caller can
/// tell "use another engine" from "this file is broken".
const EXIT_UNSUPPORTED: i32 = 3;

/// `docx-to-pdf` exit status when the input cannot be read or parsed.
const EXIT_FAILED: i32 = 2;

/// Why a conversion produced no PDF.
enum ConvertError {
    /// `strict` was asked and the document holds content Oxi does not draw.
    Unsupported(Vec<oxidocs_core::ir::UnsupportedContent>),
    /// The input could not be read or parsed, or the output not written.
    Failed(String),
}

/// What a conversion wrote.
struct Converted {
    bytes: usize,
    pages: usize,
}

/// Convert `input` to the PDF at `output`. Shared by `docx-to-pdf` and every
/// request `serve` handles, so neither exits the process on a bad document.
fn convert(input: &str, output: &str, strict: bool) -> Result<Converted, ConvertError> {
    let pdf_doc = lay_out(input, strict)?;
    let pdf_bytes = oxipdf_core::write_pdf(&pdf_doc);

    fs::write(output, &pdf_bytes).map_err(|e| ConvertError::Failed(format!("write {output}: {e}")))?;
    Ok(Converted { bytes: pdf_bytes.len(), pages: pdf_doc.pages.len() })
}

/// Draw page `page` (1-based) of `input` to the PNG at `output`, from the same
/// pages `convert` writes as PDF: a thumbnail shows what the PDF shows.
fn convert_png(input: &str, output: &str, strict: bool, page: usize, size: raster::Size) -> Result<Converted, ConvertError> {
    let pdf_doc = lay_out(input, strict)?;
    let pages = pdf_doc.pages.len();
    if page == 0 || page > pages {
        return Err(ConvertError::Failed(format!("{input} has {pages} page(s), no page {page}")));
    }
    let stand_in = stand_in_face_bytes();
    let pixmap = raster::render_page(&pdf_doc, page - 1, size, stand_in.as_deref())
        .ok_or_else(|| ConvertError::Failed(format!("cannot draw page {page} of {input} at {size:?}")))?;
    let png = raster::encode_png(&pixmap).map_err(|e| ConvertError::Failed(format!("encode {output}: {e}")))?;
    fs::write(output, &png).map_err(|e| ConvertError::Failed(format!("write {output}: {e}")))?;
    Ok(Converted { bytes: png.len(), pages })
}

/// A sans face for text the writer left unembedded (the PDF names a standard
/// font there and a viewer substitutes; the rasteriser needs real outlines).
fn stand_in_face_bytes() -> Option<Vec<u8>> {
    ["Liberation Sans", "Arial", "Carlito", "DejaVu Sans"]
        .iter()
        .find_map(|family| resolve_face_path(family, false, false))
        .and_then(|path| fs::read(path).ok())
}

/// Parse and lay out `input` into the pages `write_pdf` writes, refusing it
/// under `strict` when it holds content Oxi does not draw.
fn lay_out(input: &str, strict: bool) -> Result<PdfDocument, ConvertError> {
    let data = fs::read(input).map_err(|e| ConvertError::Failed(format!("read {input}: {e}")))?;
    let doc = oxidocs_core::parse_docx(&data).map_err(|e| ConvertError::Failed(format!("parse {input}: {e}")))?;

    note!("Parsed: {} section(s), {} blocks total",
        doc.pages.len(),
        doc.pages.iter().map(|p| p.blocks.len()).sum::<usize>());

    if !doc.unsupported.is_empty() {
        if strict {
            // Checked before layout, so a caller falling back loses no time.
            return Err(ConvertError::Unsupported(doc.unsupported));
        }
        for u in &doc.unsupported {
            eprintln!("Warning: {} x{} in {} is not drawn - the PDF is missing it", u.kind, u.count, u.part);
        }
    }

    let (pdf_doc, hf_parts_shown) = doc_to_pdf(&doc);

    if let Some((coverage, chars)) = drawn_coverage(&data, &pdf_doc, &hf_parts_shown) {
        note!("Drawn text coverage: {coverage:.4} of {chars} characters");
        if strict && coverage < MIN_DRAWN_COVERAGE {
            // Nothing in the census named it, but text was lost all the same,
            // by the parser or the layout: a table's cells, a text box, half
            // a form. Answered like any other content Oxi does not draw, so
            // the caller can fall back.
            return Err(ConvertError::Unsupported(vec![oxidocs_core::ir::UnsupportedContent {
                kind: format!("undrawn-text ({:.0}% drawn)", coverage * 100.0),
                part: "layout".to_string(),
                count: 1,
            }]));
        }
    }

    Ok(pdf_doc)
}

/// Below this share of its own text drawn, `--strict` refuses a document.
const MIN_DRAWN_COVERAGE: f64 = 0.90;

/// The share of the document's own text that the pages draw, counting each
/// character as often as it occurs (case folded, whitespace left out), and
/// how many characters that was out of. `None` for a document with too
/// little text to judge or one whose parts cannot be read.
///
/// "Its own text" is read off the package, not the parsed document: every
/// `w:t` of the main story, the header and footer parts some page shows
/// (`hf_parts_shown`, the variants the layout picked page by page: a
/// section's references to parts its pages never use count for nothing),
/// footnotes and endnotes, less the `mc:Fallback` copies that repeat a text
/// box and the runs a `w:vanish` or a hiding style takes out. Text the
/// parser never picked up is exactly what a check on the parsed document
/// cannot see, and on a 4,988-document sample most lost text was lost there.
/// Symbol-font codes are left out because they are drawn as the Unicode
/// symbol they stand for.
fn drawn_coverage(data: &[u8], pdf: &PdfDocument, hf_parts_shown: &std::collections::BTreeSet<String>) -> Option<(f64, usize)> {
    fn counted(c: char) -> bool {
        !c.is_whitespace() && !c.is_control() && !('\u{E000}'..='\u{F8FF}').contains(&c)
    }
    let mut package = oxidocs_common::archive::OoxmlArchive::new(data).ok()?;
    package.read_part("word/document.xml").ok()?;
    let mut stories = vec!["word/document.xml".to_string()];
    stories.extend(hf_parts_shown.iter().cloned());
    stories.extend(
        ["word/footnotes.xml", "word/endnotes.xml"]
            .into_iter()
            .filter(|name| package.file_names().iter().any(|n| n == name))
            .map(str::to_string),
    );
    let hidden = package.read_part("word/styles.xml").map(|xml| hidden_styles(&xml)).unwrap_or_default();
    let mut wanted: HashMap<char, usize> = HashMap::new();
    for name in stories {
        let Ok(xml) = package.read_part(&name) else { continue };
        for text in story_text(&xml, &hidden) {
            for c in text.chars().flat_map(char::to_lowercase).filter(|c| counted(*c)) {
                *wanted.entry(c).or_default() += 1;
            }
        }
    }
    let total: usize = wanted.values().sum();
    if total < 50 {
        return None;
    }
    let mut drawn: HashMap<char, usize> = HashMap::new();
    for span in pdf.pages.iter().flat_map(|page| &page.contents).filter_map(|e| e.text_span()) {
        for c in span.text.chars().flat_map(char::to_lowercase) {
            *drawn.entry(c).or_default() += 1;
        }
    }
    let found: usize = wanted.iter().map(|(c, n)| (*n).min(drawn.get(c).copied().unwrap_or(0))).sum();
    Some((found as f64 / total as f64, total))
}

/// The text of every visible run in one WordprocessingML story part: the
/// `w:t` contents of each `w:r` that neither its own `w:vanish`, its
/// character style nor its paragraph's style hides (`hidden` holds the style
/// IDs that hide), with `mc:Fallback` blocks (a second copy of what their
/// `mc:Choice` holds) skipped, and so are the cached results of the fields
/// the engine recomputes (`recomputed_field`). A scan rather than a parse: it only has to
/// count characters, and it must see what the parser may have missed.
fn story_text(xml: &str, hidden: &std::collections::HashSet<String>) -> Vec<String> {
    let mut visible = String::with_capacity(xml.len());
    let mut rest = xml;
    while let Some(start) = rest.find("<mc:Fallback") {
        visible.push_str(&rest[..start]);
        rest = match rest[start..].find("</mc:Fallback>") {
            Some(end) => &rest[start + end + "</mc:Fallback>".len()..],
            None => "",
        };
    }
    visible.push_str(rest);

    let mut texts = Vec::new();
    let mut paragraph_hidden = false;
    // Open complex fields: (instruction so far, past its `separate`).
    let mut fields: Vec<(String, bool)> = Vec::new();
    let mut pos = 0;
    while let Some(offset) = visible[pos..].find("<w:") {
        let at = pos + offset;
        let tail = &visible[at + 3..];
        let tag_end = tail.find('>').map_or(visible.len(), |e| at + 3 + e + 1);
        if tail.starts_with("p>") || tail.starts_with("p ") {
            paragraph_hidden = false;
            pos = tag_end;
        } else if tail.starts_with("pStyle ") {
            paragraph_hidden = attr_value(&visible[at..tag_end], "w:val").is_some_and(|id| hidden.contains(id));
            pos = tag_end;
        } else if tail.starts_with("fldSimple ") {
            let tag = &visible[at..tag_end];
            pos = tag_end;
            if !tag.ends_with("/>") && attr_value(tag, "w:instr").is_some_and(|i| recomputed_field(&unescape_xml(i))) {
                pos = visible[pos..].find("</w:fldSimple>").map_or(visible.len(), |e| pos + e);
            }
        } else if tail.starts_with("r>") || tail.starts_with("r ") {
            // <w:r> or <w:r ...>, not <w:rPr>, <w:rFonts>, <w:rStyle>...
            let end = tail.find("</w:r>").map_or(visible.len(), |e| at + 3 + e);
            let run = &visible[at..end];
            pos = end;
            let styled = run.find("<w:rStyle ").is_some_and(|i| {
                let tag = &run[i..run[i..].find('>').map_or(run.len(), |e| i + e)];
                attr_value(tag, "w:val").is_some_and(|id| hidden.contains(id))
            });
            let run_hidden = vanish_in(run).unwrap_or(styled || paragraph_hidden);
            let mut body = run;
            while let Some(t) = body.find("<w:") {
                let tail = &body[t + 3..];
                let Some(open_end) = tail.find('>') else { break };
                let tag = &tail[..open_end];
                let content = &tail[open_end + 1..];
                body = content;
                if tag.starts_with("fldChar ") {
                    match attr_value(tag, "w:fldCharType") {
                        Some("begin") => fields.push((String::new(), false)),
                        Some("separate") => if let Some(f) = fields.last_mut() { f.1 = true },
                        Some("end") => { fields.pop(); }
                        _ => {}
                    }
                    continue;
                }
                let is_text = tag == "t" || tag.starts_with("t ");
                let is_instr = tag == "instrText" || tag.starts_with("instrText ");
                if !(is_text || is_instr) || tag.ends_with('/') {
                    continue;
                }
                let close = content.find("</w:").unwrap_or(content.len());
                let text = unescape_xml(&content[..close]);
                body = &content[close..];
                if is_instr {
                    if let Some(f) = fields.last_mut().filter(|f| !f.1) { f.0.push_str(&text) }
                } else if !run_hidden && !fields.iter().any(|(instr, result)| *result && recomputed_field(instr)) {
                    texts.push(text);
                }
            }
        } else {
            pos = at + 3;
        }
    }
    texts
}

/// A field whose result the engine works out itself (page numbers, a
/// running style reference), so the cached result in the package is not
/// what the page shows.
fn recomputed_field(instruction: &str) -> bool {
    let name = instruction.split_whitespace().next().unwrap_or("").to_ascii_uppercase();
    matches!(name.as_str(), "PAGE" | "NUMPAGES" | "SECTIONPAGES" | "SECTION" | "STYLEREF")
}

/// What the first `w:vanish` in `xml` says, if there is one: hidden unless
/// its `w:val` is 0 or false.
fn vanish_in(xml: &str) -> Option<bool> {
    toggle_in(xml, "<w:vanish")
}

/// What the first on/off element opening with `open` (`<w:titlePg`, say) in
/// `xml` says, if there is one: on unless its `w:val` is 0, false or off.
fn toggle_in(xml: &str, open: &str) -> Option<bool> {
    let i = xml.match_indices(open).map(|(i, _)| i).find(|i| {
        xml[i + open.len()..].starts_with(['/', '>', ' '])
    })?;
    let tag = &xml[i..xml[i..].find('>').map_or(xml.len(), |e| i + e)];
    Some(!["0", "false", "off"].iter().any(|off| tag.contains(&format!("w:val=\"{off}\""))))
}

/// The value of attribute `name` in the tag text `tag`.
fn attr_value<'a>(tag: &'a str, name: &str) -> Option<&'a str> {
    let start = tag.find(&format!("{name}=\""))? + name.len() + 2;
    let len = tag[start..].find('"')?;
    Some(&tag[start..start + len])
}

/// IDs of the styles in `styles.xml` that hide their text: by their own
/// `w:vanish`, or by one inherited through `w:basedOn`. A style's `w:rPr` is
/// its run formatting, so any `w:vanish` in it hides the text it styles --
/// "Note to specifier" paragraphs of a specification template are the usual
/// case, printed by neither Word nor any other engine.
fn hidden_styles(xml: &str) -> std::collections::HashSet<String> {
    let mut styles: HashMap<String, (Option<bool>, Option<String>)> = HashMap::new();
    let mut rest = xml;
    while let Some(start) = rest.find("<w:style ") {
        let tail = &rest[start..];
        let end = tail.find("</w:style>").unwrap_or(tail.len());
        let style = &tail[..end];
        rest = &tail[end..];
        let open = &style[..style.find('>').unwrap_or(style.len())];
        let Some(id) = attr_value(open, "w:styleId") else { continue };
        let based_on = style.find("<w:basedOn ").and_then(|i| attr_value(&style[i..], "w:val")).map(str::to_string);
        styles.insert(id.to_string(), (vanish_in(style), based_on));
    }
    styles
        .keys()
        .filter(|id| {
            let mut current = id.as_str();
            for _ in 0..16 {
                match styles.get(current) {
                    Some((Some(hides), _)) => return *hides,
                    Some((None, Some(parent))) => current = parent,
                    _ => return false,
                }
            }
            false
        })
        .cloned()
        .collect()
}

/// `text` with XML's five named entities and numeric character references
/// replaced by the characters they stand for.
fn unescape_xml(text: &str) -> String {
    let mut out = String::with_capacity(text.len());
    let mut rest = text;
    while let Some(amp) = rest.find('&') {
        out.push_str(&rest[..amp]);
        let tail = &rest[amp..];
        let Some(semi) = tail.find(';') else {
            out.push_str(tail);
            return out;
        };
        let entity = &tail[1..semi];
        let decoded = match entity {
            "lt" => Some('<'),
            "gt" => Some('>'),
            "amp" => Some('&'),
            "quot" => Some('"'),
            "apos" => Some('\''),
            _ => entity
                .strip_prefix("#x")
                .map(|hex| u32::from_str_radix(hex, 16))
                .or_else(|| entity.strip_prefix('#').map(|dec| dec.parse::<u32>()))
                .and_then(|n| n.ok())
                .and_then(char::from_u32),
        };
        match decoded {
            Some(c) => out.push(c),
            None => out.push_str(&tail[..=semi]),
        }
        rest = &tail[semi + 1..];
    }
    out.push_str(rest);
    out
}

fn docx_to_pdf(input: &str, output: &str, strict: bool) {
    finish(input, output, convert(input, output, strict));
}

/// Report a one-shot conversion and exit with its status.
fn finish(input: &str, output: &str, result: Result<Converted, ConvertError>) {
    match result {
        Ok(_) if quiet() => {}
        Ok(c) => println!("Converted {} -> {} ({} bytes, {} pages)", input, output, c.bytes, c.pages),
        Err(ConvertError::Unsupported(u)) => {
            eprintln!("{}", serde_json::json!({ "unsupported": u }));
            std::process::exit(EXIT_UNSUPPORTED);
        }
        Err(ConvertError::Failed(why)) => {
            eprintln!("Error: {why}");
            std::process::exit(EXIT_FAILED);
        }
    }
}

/// Convert documents for as long as stdin stays open, one request per line:
/// `{"input": "a.docx", "output": "a.pdf", "strict": true, "id": 7}` (strict
/// and id optional). One answer per line goes to stdout as each finishes,
/// carrying the request's id:
/// `{"id":7,"ok":true,"pages":3,"bytes":51234,"ms":412.5}`, or `"ok":false`
/// with `"exit"` -- the status `docx-to-pdf` would have exited with (3 with
/// `"unsupported"`, 2 or 101 with `"error"`) -- so a caller treats both the
/// same way. A panic fails its own request and the server goes on.
///
/// Fonts, metrics and the font index load once and stay warm, which is the
/// point: a process per document pays for them every time. `threads` workers
/// convert side by side.
fn serve(threads: usize) {
    use std::io::{BufRead, Write};
    use std::sync::{mpsc, Arc, Mutex};

    let (jobs_tx, jobs_rx) = mpsc::channel::<String>();
    let jobs_rx = Arc::new(Mutex::new(jobs_rx));
    let out = Arc::new(Mutex::new(std::io::stdout()));
    let workers: Vec<_> = (0..threads)
        .map(|_| {
            let jobs_rx = Arc::clone(&jobs_rx);
            let out = Arc::clone(&out);
            std::thread::spawn(move || loop {
                let Ok(line) = jobs_rx.lock().expect("job queue").recv() else { break };
                let answer = serve_one(&line);
                let mut out = out.lock().expect("stdout");
                let _ = writeln!(out, "{answer}");
                let _ = out.flush();
            })
        })
        .collect();
    for line in std::io::stdin().lock().lines().map_while(Result::ok) {
        if !line.trim().is_empty() && jobs_tx.send(line).is_err() {
            break;
        }
    }
    drop(jobs_tx);
    for w in workers {
        let _ = w.join();
    }
}

fn serve_one(line: &str) -> serde_json::Value {
    use serde_json::json;
    let req: serde_json::Value = match serde_json::from_str(line) {
        Ok(v) => v,
        Err(e) => return json!({ "ok": false, "exit": EXIT_FAILED, "error": format!("request: {e}") }),
    };
    let id = req.get("id").cloned().unwrap_or(serde_json::Value::Null);
    let (Some(input), Some(output)) = (req["input"].as_str(), req["output"].as_str()) else {
        return json!({ "id": id, "ok": false, "exit": EXIT_FAILED, "error": "request needs input and output" });
    };
    let strict = req["strict"].as_bool().unwrap_or(false);
    let png = output.to_ascii_lowercase().ends_with(".png");
    let number = |key: &str| req[key].as_u64().filter(|n| *n > 0 && *n <= u32::MAX as u64).map(|n| n as u32);
    let page = number("page").unwrap_or(1) as usize;
    let size = raster::Size { width: number("width"), height: number("height") };
    let start = std::time::Instant::now();
    let result = std::panic::catch_unwind(|| {
        if png {
            convert_png(input, output, strict, page, size)
        } else {
            convert(input, output, strict)
        }
    });
    let ms = start.elapsed().as_secs_f64() * 1000.0;
    match result {
        Ok(Ok(c)) => json!({ "id": id, "ok": true, "pages": c.pages, "bytes": c.bytes, "ms": ms }),
        Ok(Err(ConvertError::Unsupported(u))) => json!({ "id": id, "ok": false, "exit": EXIT_UNSUPPORTED, "unsupported": u, "ms": ms }),
        Ok(Err(ConvertError::Failed(why))) => json!({ "id": id, "ok": false, "exit": EXIT_FAILED, "error": why, "ms": ms }),
        Err(panic) => {
            let why = panic
                .downcast_ref::<String>()
                .cloned()
                .or_else(|| panic.downcast_ref::<&str>().map(|s| s.to_string()))
                .unwrap_or_else(|| "panic".into());
            json!({ "id": id, "ok": false, "exit": 101, "error": why, "ms": ms })
        }
    }
}

// --- Direct Document IR → PDF conversion (no LayoutEngine) ---

struct PdfBuilder {
    pages: Vec<Page>,
    current_contents: Vec<ContentElement>,
    page_width: f64,
    page_height: f64,
    margin_top: f64,
    margin_bottom: f64,
    margin_left: f64,
    margin_right: f64,
    cursor_y: f64,
    /// Unicode → width in 1/1000 em units (from embedded font).
    /// Used for accurate text width estimation.
    font_widths: HashMap<u32, u16>,
}

impl PdfBuilder {
    /// Create a new PdfBuilder.
    /// Coordinate system: y=0 at TOP of page (top-down, matching writer's expectation).
    /// cursor_y starts at margin_top and increases downward.
    fn new(width: f64, height: f64, margin: (f64, f64, f64, f64), font_widths: HashMap<u32, u16>) -> Self {
        Self {
            pages: Vec::new(),
            current_contents: Vec::new(),
            page_width: width,
            page_height: height,
            margin_top: margin.0,
            margin_bottom: margin.1,
            margin_left: margin.2,
            margin_right: margin.3,
            cursor_y: margin.0,
            font_widths,
        }
    }

    /// Estimate text width using actual font metrics when available.
    fn text_width(&self, text: &str, font_size: f64) -> f64 {
        if self.font_widths.is_empty() {
            return estimate_text_width(text, font_size);
        }
        let mut w = 0.0;
        for ch in text.chars() {
            if let Some(&glyph_w) = self.font_widths.get(&(ch as u32)) {
                // glyph_w is in 1/1000 em units
                w += font_size * (glyph_w as f64 / 1000.0);
            } else {
                w += char_width(ch, font_size);
            }
        }
        w
    }

    /// Wrap text using actual font metrics.
    fn wrap_text_accurate(&self, text: &str, font_size: f64, max_width: f64) -> Vec<String> {
        if self.font_widths.is_empty() {
            return wrap_text(text, font_size, max_width);
        }
        let mut lines = Vec::new();
        let mut current_line = String::new();
        let mut current_width = 0.0;

        for ch in text.chars() {
            let ch_w = if let Some(&glyph_w) = self.font_widths.get(&(ch as u32)) {
                font_size * (glyph_w as f64 / 1000.0)
            } else {
                char_width(ch, font_size)
            };

            if current_width + ch_w > max_width && !current_line.is_empty() {
                lines.push(current_line);
                current_line = String::new();
                current_width = 0.0;
            }

            current_line.push(ch);
            current_width += ch_w;
        }

        if !current_line.is_empty() {
            lines.push(current_line);
        }
        if lines.is_empty() {
            lines.push(String::new());
        }
        lines
    }

    fn content_width(&self) -> f64 {
        self.page_width - self.margin_left - self.margin_right
    }

    fn max_y(&self) -> f64 {
        self.page_height - self.margin_bottom
    }

    fn needs_page_break(&self, needed_height: f64) -> bool {
        self.cursor_y + needed_height > self.max_y()
    }

    fn new_page(&mut self) {
        let page = Page {
            width: self.page_width,
            height: self.page_height,
            media_box: Rectangle {
                llx: 0.0, lly: 0.0,
                urx: self.page_width, ury: self.page_height,
            },
            crop_box: None,
            contents: std::mem::take(&mut self.current_contents),
            rotation: 0,
        };
        self.pages.push(page);
        self.cursor_y = self.margin_top;
    }

    fn finish(mut self) -> Vec<Page> {
        if !self.current_contents.is_empty() {
            let page = Page {
                width: self.page_width,
                height: self.page_height,
                media_box: Rectangle {
                    llx: 0.0, lly: 0.0,
                    urx: self.page_width, ury: self.page_height,
                },
                crop_box: None,
                contents: std::mem::take(&mut self.current_contents),
                rotation: 0,
            };
            self.pages.push(page);
        }
        self.pages
    }

    /// Add text at (x, y) where y is in top-down coordinates.
    fn add_text(&mut self, x: f64, y: f64, text: String, font_name: String, font_size: f64, color: Color) {
        self.current_contents.push(ContentElement::Text(TextSpan {
            x, y, text, font_name, font_size, fill_color: color, character_spacing: 0.0, stroke_width: 0.0, skew: 0.0, horizontal_scale: 100.0, stroke_color: None,
        }));
    }

    fn add_line(&mut self, x1: f64, y1: f64, x2: f64, y2: f64, width: f64, color: Color) {
        self.current_contents.push(ContentElement::Path(PathData {
            operations: vec![
                PathOp::MoveTo(x1, y1),
                PathOp::LineTo(x2, y2),
            ],
            stroke: Some(StrokeStyle {
                color,
                width,
                line_cap: LineCap::Butt,
                line_join: LineJoin::Miter,
            }),
            fill: None,
        }));
    }

    fn add_rect_fill(&mut self, x: f64, y: f64, w: f64, h: f64, color: Color) {
        self.current_contents.push(ContentElement::Path(PathData {
            operations: vec![
                PathOp::MoveTo(x, y),
                PathOp::LineTo(x + w, y),
                PathOp::LineTo(x + w, y + h),
                PathOp::LineTo(x, y + h),
                PathOp::ClosePath,
            ],
            stroke: None,
            fill: Some(color),
        }));
    }

    fn add_rounded_rect(&mut self, x: f64, y: f64, w: f64, h: f64, r: f64, fill: Option<Color>, stroke: Option<(f64, Color)>) {
        // Bézier control point factor for circular arcs: 4*(sqrt(2)-1)/3 ≈ 0.5523
        let k = 0.5523 * r;
        let ops = vec![
            PathOp::MoveTo(x + r, y),
            PathOp::LineTo(x + w - r, y),
            PathOp::CurveTo(x + w - r + k, y, x + w, y + r - k, x + w, y + r),
            PathOp::LineTo(x + w, y + h - r),
            PathOp::CurveTo(x + w, y + h - r + k, x + w - r + k, y + h, x + w - r, y + h),
            PathOp::LineTo(x + r, y + h),
            PathOp::CurveTo(x + r - k, y + h, x, y + h - r + k, x, y + h - r),
            PathOp::LineTo(x, y + r),
            PathOp::CurveTo(x, y + r - k, x + r - k, y, x + r, y),
            PathOp::ClosePath,
        ];
        self.current_contents.push(ContentElement::Path(PathData {
            operations: ops,
            stroke: stroke.map(|(sw, sc)| StrokeStyle {
                color: sc, width: sw, line_cap: LineCap::Butt, line_join: LineJoin::Round,
            }),
            fill,
        }));
    }

    fn add_rect_stroke(&mut self, x: f64, y: f64, w: f64, h: f64, line_w: f64, color: Color) {
        self.current_contents.push(ContentElement::Path(PathData {
            operations: vec![
                PathOp::MoveTo(x, y),
                PathOp::LineTo(x + w, y),
                PathOp::LineTo(x + w, y + h),
                PathOp::LineTo(x, y + h),
                PathOp::ClosePath,
            ],
            stroke: Some(StrokeStyle {
                color,
                width: line_w,
                line_cap: LineCap::Butt,
                line_join: LineJoin::Miter,
            }),
            fill: None,
        }));
    }
}

/// Render header/footer blocks into a page at specified y position.
fn render_header_footer_blocks(
    builder: &mut PdfBuilder,
    blocks: &[Block],
    y_start: f64,
    registry: &FontMetricsRegistry,
    doc_default_font: Option<&str>,
) {
    let saved_y = builder.cursor_y;
    builder.cursor_y = y_start;
    let ml = builder.margin_left;
    let cw = builder.content_width();

    for block in blocks {
        if let Block::Paragraph(para) = block {
            let text: String = para.runs.iter().map(|r| r.text.as_str()).collect();
            if text.trim().is_empty() { continue; }

            let (font_size, font_family, bold) = para_font_props(para);
            let font_name = resolve_font(font_family, bold);
            let color = para_color(para);
            let text_w = builder.text_width(&text, font_size);
            let x = match para.alignment {
                Alignment::Right => ml + cw - text_w,
                Alignment::Center => ml + (cw - text_w) / 2.0,
                _ => ml,
            };
            let lh = compute_line_height(
                registry, font_family, font_size as f32,
                para.style.line_spacing, para.style.line_spacing_rule.as_deref(),
                doc_default_font, 0.0,
            );
            builder.cursor_y += lh;
            add_text_split(builder, x, builder.cursor_y, text, font_name, font_size, color);
        }
    }
    builder.cursor_y = saved_y;
}

/// Render paragraph shading (background fill).
fn render_paragraph_shading(
    builder: &mut PdfBuilder,
    shading: &str,
    x: f64, y: f64, w: f64, h: f64,
) {
    if let Some(color) = parse_hex_color(shading) {
        builder.add_rect_fill(x, y - 1.0, w, h + 2.0, color);
    }
}

/// Render paragraph borders.
fn render_paragraph_borders(
    builder: &mut PdfBuilder,
    borders: &oxidocs_core::ir::ParagraphBorders,
    x: f64, y: f64, w: f64, h: f64,
) {
    let draw_border = |builder: &mut PdfBuilder, border: &Option<oxidocs_core::ir::BorderDef>, x1: f64, y1: f64, x2: f64, y2: f64| {
        if let Some(bd) = border {
            let bw = if bd.width > 0.0 { bd.width as f64 } else { 0.5 };
            let color = bd.color.as_deref()
                .and_then(parse_hex_color)
                .unwrap_or(Color::Gray(0.0));
            builder.add_line(x1, y1, x2, y2, bw, color);
        }
    };
    draw_border(builder, &borders.top, x, y, x + w, y);
    draw_border(builder, &borders.bottom, x, y + h, x + w, y + h);
    draw_border(builder, &borders.left, x, y, x, y + h);
    draw_border(builder, &borders.right, x + w, y, x + w, y + h);
}

/// Resolve TextBox absolute position (matches layout/mod.rs resolve_textbox_position).
fn resolve_textbox_position(
    text_box: &TextBox,
    section: &oxidocs_core::ir::Page,
    block_y_positions: &[f64],
) -> (f64, f64) {
    let pos = match &text_box.position {
        Some(p) => p,
        None => return (section.margin.left as f64, section.margin.top as f64),
    };

    let page_w = section.size.width as f64;
    let page_h = section.size.height as f64;
    let ml = section.margin.left as f64;
    let mr = section.margin.right as f64;
    let mt = section.margin.top as f64;
    let mb = section.margin.bottom as f64;
    let content_w = page_w - ml - mr;
    let tb_w = text_box.width as f64;
    let tb_h = text_box.height as f64;

    // Horizontal position
    let abs_x = if let Some(ref align) = pos.h_align {
        // Determine reference frame from h_relative
        let (ref_left, ref_w) = match pos.h_relative.as_deref() {
            Some("page") => (0.0, page_w),
            Some("margin") | Some("column") | None => (ml, content_w),
            Some("leftMarginArea") => (0.0, ml),
            Some("rightMarginArea") => (page_w - mr, mr),
            _ => (ml, content_w),
        };
        match align.as_str() {
            "left" => ref_left,
            "center" => ref_left + (ref_w - tb_w) / 2.0,
            "right" => ref_left + ref_w - tb_w,
            _ => ref_left + pos.x as f64,
        }
    } else {
        let offset = pos.x as f64;
        match pos.h_relative.as_deref() {
            Some("page") => offset,
            Some("margin") | Some("column") | None => ml + offset,
            Some("leftMarginArea") => offset,
            Some("rightMarginArea") => (page_w - mr) + offset,
            _ => ml + offset,
        }
    };

    // Vertical position
    let abs_y = if let Some(ref align) = pos.v_align {
        let (ref_top, ref_h) = match pos.v_relative.as_deref() {
            Some("page") => (0.0, page_h),
            Some("margin") | None => (mt, page_h - mt - mb),
            _ => (mt, page_h - mt - mb),
        };
        match align.as_str() {
            "top" => ref_top,
            "center" => ref_top + (ref_h - tb_h) / 2.0,
            "bottom" => ref_top + ref_h - tb_h,
            _ => ref_top + pos.y as f64,
        }
    } else {
        let offset = pos.y as f64;
        match pos.v_relative.as_deref() {
            Some("page") => offset,
            Some("margin") | None => mt + offset,
            Some("paragraph") | Some("line") => {
                let anchor_y = block_y_positions
                    .get(text_box.anchor_block_index)
                    .copied()
                    .unwrap_or(mt);
                anchor_y + offset
            }
            Some("topMarginArea") => offset,
            Some("bottomMarginArea") => (page_h - mb) + offset,
            _ => mt + offset,
        }
    };

    (abs_x, abs_y)
}

/// Render a TextBox: background fill, border, and inner content blocks.
fn render_textbox(
    builder: &mut PdfBuilder,
    text_box: &TextBox,
    section: &oxidocs_core::ir::Page,
    block_y_positions: &[f64],
    registry: &FontMetricsRegistry,
    doc_default_font: Option<&str>,
    doc_default_font_size: f32,
    grid_pitch_pt: f64,
) {
    let (abs_x, abs_y) = resolve_textbox_position(text_box, section, block_y_positions);
    let tb_w = text_box.width as f64;
    let tb_h = text_box.height as f64;

    let cr = text_box.corner_radius.map(|r| r as f64).unwrap_or(0.0);
    let has_rounded = cr > 0.0;

    // Background fill + Border
    if has_rounded {
        let fill_color = text_box.fill.as_ref().and_then(|h| parse_hex_color(h));
        let stroke = if text_box.border { Some((0.5, Color::Gray(0.0))) } else { None };
        if fill_color.is_some() || stroke.is_some() {
            builder.add_rounded_rect(abs_x, abs_y, tb_w, tb_h, cr, fill_color, stroke);
        }
    } else {
        if let Some(ref fill_hex) = text_box.fill {
            if let Some(color) = parse_hex_color(fill_hex) {
                builder.add_rect_fill(abs_x, abs_y, tb_w, tb_h, color);
            }
        }
        if text_box.border {
            let bc = Color::Gray(0.0);
            let bw = 0.5;
            builder.add_line(abs_x, abs_y, abs_x + tb_w, abs_y, bw, bc);
            builder.add_line(abs_x, abs_y + tb_h, abs_x + tb_w, abs_y + tb_h, bw, bc);
            builder.add_line(abs_x, abs_y, abs_x, abs_y + tb_h, bw, bc);
            builder.add_line(abs_x + tb_w, abs_y, abs_x + tb_w, abs_y + tb_h, bw, bc);
        }
    }

    // Inner content with inset (Word default: L/R=7.2pt, T/B=3.6pt)
    let inset_lr = 7.2;
    let inset_tb = 3.6;
    let inner_x = abs_x + inset_lr;
    let inner_w = tb_w - inset_lr * 2.0;
    let saved_y = builder.cursor_y;
    builder.cursor_y = abs_y + inset_tb;

    for block in &text_box.blocks {
        match block {
            Block::Paragraph(para) => {
                let gp = if para.style.snap_to_grid { grid_pitch_pt } else { 0.0 };
                render_paragraph_styled(builder, para, ParaRole::Body, inner_x, inner_w, registry, doc_default_font, doc_default_font_size, gp);
            }
            Block::Table(table) => {
                render_table(builder, table, inner_x, inner_w, registry, doc_default_font);
            }
            _ => {}
        }
    }

    builder.cursor_y = saved_y;
}

/// Render footnotes at the bottom of the current page.
fn render_footnotes(
    builder: &mut PdfBuilder,
    footnotes: &[oxidocs_core::ir::Footnote],
    registry: &FontMetricsRegistry,
    doc_default_font: Option<&str>,
) {
    if footnotes.is_empty() { return; }
    let ml = builder.margin_left;
    let cw = builder.content_width();
    let footnote_font_size = 8.0_f64;

    // Separator line
    let sep_y = builder.max_y() - (footnotes.len() as f64 * footnote_font_size * 1.4) - 10.0;
    builder.add_line(ml, sep_y, ml + cw * 0.3, sep_y, 0.5, Color::Gray(0.3));

    let mut fy = sep_y + 4.0;
    for note in footnotes {
        let text: String = note.blocks.iter().filter_map(|b| {
            if let Block::Paragraph(p) = b {
                Some(p.runs.iter().map(|r| r.text.as_str()).collect::<String>())
            } else { None }
        }).collect::<Vec<_>>().join(" ");

        let marker = format!("{}  ", note.number);
        let full = format!("{}{}", marker, text);
        add_text_split(builder, ml, fy, full, "Helvetica".to_string(), footnote_font_size, Color::Gray(0.2));
        fy += footnote_font_size * 1.4;
    }
}

/// Compute line height following Word's layout behavior.
///
/// Line spacing rules (ECMA-376 17.3.1.13 + Word COM measurements):
/// - auto (multiple): base_height * multiplier
/// - atLeast: max(base_height, specified_pts)
/// - exact: specified_pts exactly
///
/// base_height = max(run_line_height, default_font_line_height)
/// Compute line height following Word's layout behavior.
///
/// grid_pitch_pt: document grid pitch in points (default 18pt = 360 twips).
///   Set to 0.0 for table cells (DontSnapToGridInCell).
fn compute_line_height(
    registry: &FontMetricsRegistry,
    font_family: Option<&str>,
    font_size: f32,
    line_spacing: Option<f32>,
    line_spacing_rule: Option<&str>,
    doc_default_font: Option<&str>,
    grid_pitch_pt: f64,
) -> f64 {
    // grid_pitch_pt > 0 implies snap_to_grid=true
    let snap = grid_pitch_pt > 0.01;
    compute_line_height_with_default(registry, font_family, font_size, line_spacing, line_spacing_rule, doc_default_font, None, grid_pitch_pt, snap)
}

fn compute_line_height_with_default(
    registry: &FontMetricsRegistry,
    font_family: Option<&str>,
    font_size: f32,
    line_spacing: Option<f32>,
    line_spacing_rule: Option<&str>,
    doc_default_font: Option<&str>,
    doc_default_font_size: Option<f32>,
    grid_pitch_pt: f64,
    snap_to_grid: bool,
) -> f64 {
    let run_metrics = font_family
        .map(|f| registry.get(f))
        .unwrap_or(registry.default_metrics());
    let run_base = run_metrics.word_line_height(font_size, 96.0) as f64;

    // snap_to_grid=false uses natural font height only,
    // NO default font minimum. snap_to_grid=true applies max(run, default).
    let base_height = if snap_to_grid {
        let default_metrics = doc_default_font
            .map(|f| registry.get(f))
            .unwrap_or(registry.default_metrics());
        let default_fs = doc_default_font_size.unwrap_or(10.5);
        let default_base = default_metrics.word_line_height(default_fs, 96.0) as f64;
        run_base.max(default_base)
    } else {
        run_base
    };

    let raw = match line_spacing_rule {
        Some("exact") => {
            // Exact: use the specified value in points directly, no grid snap
            return line_spacing.unwrap_or(font_size) as f64;
        }
        Some("atLeast") => {
            let min_height = line_spacing.unwrap_or(0.0) as f64;
            base_height.max(min_height)
        }
        _ => {
            // Auto/multiple: base_height * multiplier
            let mult = line_spacing.unwrap_or(0.0) as f64;
            if mult > 0.01 && mult < 10.0 {
                base_height * mult
            } else {
                base_height // single spacing (1.0)
            }
        }
    };

    // Grid snap: ceil(raw / grid_pitch) * grid_pitch
    if grid_pitch_pt > 0.01 {
        (raw / grid_pitch_pt).ceil() * grid_pitch_pt
    } else {
        raw
    }
}

/// Estimate text width. CJK chars ≈ font_size, Latin width varies by character class.
fn estimate_text_width(text: &str, font_size: f64) -> f64 {
    let mut w = 0.0;
    for ch in text.chars() {
        if is_cjk(ch) {
            w += font_size;
        } else if ch == ' ' {
            w += font_size * 0.25;
        } else if ch == '\u{3000}' {
            // Ideographic space
            w += font_size;
        } else if ch.is_ascii_uppercase() {
            w += font_size * 0.60;
        } else if ch.is_ascii_lowercase() {
            w += font_size * 0.48;
        } else if ch.is_ascii_digit() {
            w += font_size * 0.50;
        } else if ch == '.' || ch == ',' || ch == ':' || ch == ';' || ch == '!' || ch == '\'' {
            w += font_size * 0.25;
        } else if ch == '(' || ch == ')' || ch == '[' || ch == ']' {
            w += font_size * 0.30;
        } else if ch == '—' || ch == '–' || ch == '〜' {
            w += font_size * 0.60;
        } else {
            w += font_size * 0.50;
        }
    }
    w
}

fn is_cjk(ch: char) -> bool {
    let c = ch as u32;
    // CJK Unified, Hiragana, Katakana, fullwidth punctuation, etc.
    (0x3000..=0x9FFF).contains(&c)
        || (0xF900..=0xFAFF).contains(&c)
        || (0xFF00..=0xFFEF).contains(&c)
        || (0x20000..=0x2FA1F).contains(&c)
        || ch == '①' || ch == '②' || ch == '③' || ch == '④' || ch == '⑤'
        || ch == '⑥' || ch == '⑦' || ch == '⑧' || ch == '⑨' || ch == '⑩'
        || ch == '→' || ch == '◆' || ch == '■' || ch == '●' || ch == '◎'
}

/// Break text into lines that fit within max_width
fn wrap_text(text: &str, font_size: f64, max_width: f64) -> Vec<String> {
    if text.is_empty() {
        return vec![String::new()];
    }

    let mut lines = Vec::new();
    let mut current_line = String::new();
    let mut current_width = 0.0;

    for ch in text.chars() {
        let ch_width = char_width(ch, font_size);

        if current_width + ch_width > max_width && !current_line.is_empty() {
            lines.push(current_line);
            current_line = String::new();
            current_width = 0.0;
        }

        current_line.push(ch);
        current_width += ch_width;
    }

    if !current_line.is_empty() {
        lines.push(current_line);
    }

    if lines.is_empty() {
        lines.push(String::new());
    }
    lines
}

/// Per-character width estimate for line wrapping
fn char_width(ch: char, font_size: f64) -> f64 {
    if is_cjk(ch) {
        font_size
    } else if ch == ' ' {
        font_size * 0.25
    } else if ch == '\u{3000}' {
        font_size
    } else if ch.is_ascii_uppercase() {
        font_size * 0.60
    } else if ch.is_ascii_lowercase() {
        font_size * 0.48
    } else if ch.is_ascii_digit() {
        font_size * 0.50
    } else if ch == '.' || ch == ',' || ch == ':' || ch == ';' || ch == '\'' {
        font_size * 0.25
    } else if ch == '(' || ch == ')' || ch == '[' || ch == ']' {
        font_size * 0.30
    } else {
        font_size * 0.50
    }
}

/// Faces the document actually asks for, resolved to files on this machine.
/// Built once per conversion by `build_embedded_fonts`, then read by
/// `resolve_font` so the emitter and the embedder agree on one name per face.
#[derive(Debug, Default)]
struct FontPlan {
    /// (lowercased family, bold, italic) -> PDF font key that was embedded
    keys: HashMap<(String, bool, bool), String>,
}

thread_local! {
    /// The plan of the conversion running on this thread. It was a process-wide
    /// OnceLock: `oxidocs serve` converts many documents per process, so every
    /// document after the first was drawn with the FIRST document's plan --
    /// Arial runs, say, went to the OxiLatin (Carlito-width) stand-in although
    /// Liberation Sans had just been embedded for them. A conversion runs on
    /// one thread from layout to PDF, so the plan lives with the thread and is
    /// replaced by every conversion.
    static FONT_PLAN: std::cell::RefCell<FontPlan> = std::cell::RefCell::new(FontPlan::default());
}

/// Metric-compatible free substitutes, used only when the real family is
/// absent. Measured 2026-09-02 against the Microsoft originals: Carlito and
/// the three Liberation faces match their counterpart's advance width on all
/// 95 printable ASCII characters, so substituting one moves no glyph.
/// Caladea does NOT -- it differs on every ASCII character (up to 0.19 em),
/// so a Cambria document substituted this way reflows within the line. It is
/// still a closer answer than letting the viewer pick, but it is not exact.
const SUBSTITUTES: &[(&str, &str)] = &[
    ("calibri", "Carlito"),
    ("cambria", "Caladea"),
    ("arial", "Liberation Sans"),
    ("helvetica", "Liberation Sans"),
    // Helvetica Neue deliberately has no entry. Substituting Liberation Sans
    // for it was measured on the blind corpus and came out *worse* than
    // letting it fall to base-14 Helvetica — a reminder that a substitute is
    // only worth making when its metrics are known to match, not when the
    // families merely look alike.
    ("times new roman", "Liberation Serif"),
    ("times", "Liberation Serif"),
    ("courier new", "Liberation Mono"),
    ("courier", "Liberation Mono"),
    // Liberation Sans Narrow matches Arial Narrow exactly — measured
    // 2026-09-02, all 95 printable ASCII widths identical, hhea identical.
    // It is NOT in `fonts/` and must not be added there: Narrow is still
    // Liberation 1.x, which is GPLv2-with-font-exception, and this project
    // ships only MPL-2.0-compatible third-party files. Mapping to it costs
    // nothing and redistributes nothing — it is used when the machine already
    // has it, which most Linux distributions do.
    ("arial narrow", "Liberation Sans Narrow"),
];

/// The machine's font index, built once. Every path that needs a font file
/// goes through here, so what a conversion can see is decided in one place and
/// the environment (a container with no Microsoft fonts, say) is honoured
/// consistently rather than by whichever caller happened to hardcode a path.
fn font_index() -> &'static fontidx::FontIndex {
    static INDEX: std::sync::OnceLock<fontidx::FontIndex> = std::sync::OnceLock::new();
    INDEX.get_or_init(|| fontidx::FontIndex::build(&bundled_font_dirs()))
}

/// Path to the file holding `family` at this weight: the family itself, else
/// its bundled metric-compatible stand-in. TTC members past the first are
/// skipped because the subsetter addresses files, not collection entries.
fn resolve_face_path(family: &str, bold: bool, italic: bool) -> Option<String> {
    let index = font_index();
    // The `index == 0` test belongs to each candidate, not to the pair: a
    // family that only exists deeper in a collection should still let the
    // bundled stand-in be tried, rather than taking the whole lookup down.
    index
        .find(family, bold, italic)
        .filter(|face| face.index == 0)
        .or_else(|| {
            substitute_for(family)
                .and_then(|s| index.find(s, bold, italic))
                .filter(|face| face.index == 0)
        })
        .map(|face| face.path.to_string_lossy().to_string())
}

/// Split a trailing weight or slope word off a family name, as in
/// `Trebuchet MS Bold` or `Calibri Light`. Some producers write the full face
/// name where the family belongs, and the shortened name is often one we can
/// resolve or substitute when the long one is not.
///
/// Only weight and slope are stripped. Width is not: `Arial Narrow` is its own
/// family with its own advances, and turning it into `Arial` would silently
/// change every line length — the same mistake as substituting a face that
/// merely looks similar.
fn split_style_suffix(family: &str) -> Option<(String, bool, bool)> {
    const WEIGHTS: &[&str] = &[
        "Thin", "ExtraLight", "UltraLight", "Light", "Regular", "Book",
        "Medium", "Semibold", "SemiBold", "Demibold", "DemiBold", "Bold",
        "ExtraBold", "UltraBold", "Black", "Heavy",
    ];
    const SLOPES: &[&str] = &["Italic", "Oblique"];

    let mut words: Vec<&str> = family.split_whitespace().collect();
    let mut bold = false;
    let mut italic = false;
    let mut stripped = false;
    while words.len() > 1 {
        let last = words[words.len() - 1];
        if SLOPES.iter().any(|s| s.eq_ignore_ascii_case(last)) {
            italic = true;
        } else if WEIGHTS.iter().any(|w| w.eq_ignore_ascii_case(last)) {
            // "Bold" and the heavier names are a weight we can ask for; the
            // lighter ones just are not this family's regular face, and asking
            // for regular is the closest we can honestly get.
            if last.eq_ignore_ascii_case("Bold")
                || last.eq_ignore_ascii_case("Black")
                || last.eq_ignore_ascii_case("Heavy")
            {
                bold = true;
            }
        } else {
            break;
        }
        words.pop();
        stripped = true;
    }
    if !stripped {
        return None;
    }
    Some((words.join(" "), bold, italic))
}

fn substitute_for(family: &str) -> Option<&'static str> {
    let key = family.trim().to_lowercase();
    SUBSTITUTES
        .iter()
        .find(|(name, _)| *name == key)
        .map(|(_, sub)| *sub)
}

/// PDF font key for one face. Derived from the family name alone so the
/// emitter can name a face before the plan is built, and so two runs asking
/// for the same face always land on the same key.
fn font_key(family: &str, bold: bool, italic: bool) -> String {
    let mut key = String::from("F");
    for ch in family.chars() {
        if ch.is_ascii_alphanumeric() {
            key.push(ch);
        }
    }
    if bold {
        key.push_str("_B");
    }
    if italic {
        key.push_str("_I");
    }
    key
}

/// Directories holding the fonts shipped with Oxi: `fonts/` beside the
/// executable for an installed build, and the source tree's own `fonts/` so a
/// `target/release` binary run from a checkout finds them too.
///
/// The files themselves belong to `oxidocs-fonts`, which is also where the
/// browser build gets them; it answers with its own folder only while the
/// running binary is still inside the tree that folder belongs to.
fn bundled_font_dirs() -> Vec<std::path::PathBuf> {
    let mut dirs = Vec::new();
    let exe = std::env::current_exe().ok();
    if let Some(dir) = exe.as_ref().and_then(|e| e.parent()) {
        dirs.push(dir.join("fonts"));
    }
    if let Some(exe) = exe.as_ref() {
        dirs.extend(oxidocs_fonts::source_dirs(exe));
    }
    // Installed with `cargo install`, neither of those exists: only the binary
    // is copied, and the faces it needs are compiled into it. So they are
    // written out once, next to the binary, and found by the first path on
    // every run after that. Next to the binary is often not writable -- a
    // container runs /usr/bin/oxidocs as an unprivileged user -- so the user's
    // cache and then the temporary directory are tried after it. Without
    // this, such an install ran with none of its own faces at all.
    if dirs.iter().all(|dir| !dir.is_dir()) {
        let cache = std::env::var_os("XDG_CACHE_HOME")
            .map(std::path::PathBuf::from)
            .or_else(|| std::env::var_os("HOME").map(|home| std::path::PathBuf::from(home).join(".cache")))
            .map(|dir| dir.join("oxidocs").join(concat!("fonts-", env!("CARGO_PKG_VERSION"))));
        let candidates = exe.as_ref().and_then(|e| e.parent()).map(|dir| dir.join("fonts")).into_iter()
            .chain(cache)
            .chain(std::iter::once(std::env::temp_dir().join(concat!("oxidocs-fonts-", env!("CARGO_PKG_VERSION")))));
        for candidate in candidates {
            if let Some(written) = lay_out_fonts(&candidate) {
                dirs.insert(0, written);
                break;
            }
        }
    }
    dirs
}

/// One shipped face's file, from the first of `bundled_font_dirs` that holds it.
fn bundled_face_file(file: &str) -> Option<std::path::PathBuf> {
    bundled_font_dirs()
        .into_iter()
        .map(|dir| dir.join(file))
        .find(|path| path.is_file())
}

/// Write the compiled-in faces into `where_to`, and say where they landed.
///
/// Nothing is overwritten: a face already there is somebody's deliberate copy,
/// or one of these written on an earlier run. `None` when the directory cannot
/// be made — an install under a read-only prefix, say — and the caller then
/// carries on with whatever the system provides.
fn lay_out_fonts(where_to: &std::path::Path) -> Option<std::path::PathBuf> {
    std::fs::create_dir_all(where_to).ok()?;
    for (_, file, data) in oxidocs_fonts::faces() {
        let at = where_to.join(file);
        if at.exists() {
            continue;
        }
        std::fs::write(&at, data).ok()?;
    }
    Some(where_to.to_path_buf())
}

fn resolve_font(family: Option<&str>, bold: bool) -> String {
    resolve_font_styled(family, bold, false)
}

/// a:srcRect crop (percent insets top, right, bottom, left; a negative inset adds
/// transparent margin) and a quarter / half turn of the pixels when the element box is
/// already the turned one -- the DirectWrite renderer's render_image rules.
fn pdf_image_crop_turn(
    img: image::DynamicImage,
    crop: Option<(f32, f32, f32, f32)>,
    rotation: f32,
    w_pt: f32,
    h_pt: f32,
) -> image::DynamicImage {
    let img = match crop {
        Some((t, r, b, l)) if t > 0.0 || r > 0.0 || b > 0.0 || l > 0.0 => {
            let (pw, ph) = (img.width(), img.height());
            let x0 = (pw as f32 * l / 100.0).round().max(0.0) as u32;
            let y0 = (ph as f32 * t / 100.0).round().max(0.0) as u32;
            let cw = (pw as f32 * (100.0 - l - r) / 100.0).round().max(1.0) as u32;
            let ch = (ph as f32 * (100.0 - t - b) / 100.0).round().max(1.0) as u32;
            if x0.saturating_add(cw) <= pw && y0.saturating_add(ch) <= ph
                && t >= 0.0 && r >= 0.0 && b >= 0.0 && l >= 0.0
            {
                img.crop_imm(x0, y0, cw, ch)
            } else {
                let (fw, fh) = (pw as f32, ph as f32);
                let (sx, sy) = (fw * l / 100.0, fh * t / 100.0);
                let (sw, sh) = (fw * (100.0 - l - r) / 100.0, fh * (100.0 - t - b) / 100.0);
                let (ix0, iy0) = (sx.max(0.0), sy.max(0.0));
                let (ix1, iy1) = ((sx + sw).min(fw), (sy + sh).min(fh));
                if sw >= 1.0 && sh >= 1.0 && ix1 - ix0 >= 1.0 && iy1 - iy0 >= 1.0 {
                    let mut canvas = image::RgbaImage::new(sw.round() as u32, sh.round() as u32);
                    let part = img.crop_imm(ix0.round() as u32, iy0.round() as u32,
                        (ix1 - ix0).round() as u32, (iy1 - iy0).round() as u32).to_rgba8();
                    image::imageops::overlay(&mut canvas, &part,
                        (ix0 - sx).round() as i64, (iy0 - sy).round() as i64);
                    image::DynamicImage::ImageRgba8(canvas)
                } else {
                    img
                }
            }
        }
        _ => img,
    };
    let turn_deg = rotation.rem_euclid(360.0).round() as i32;
    let box_unturned = (turn_deg == 90 || turn_deg == 270) && {
        let (iw, ih) = (img.width() as f32, img.height() as f32);
        let ba = w_pt / h_pt.max(1e-3);
        (ba - iw / ih.max(1.0)).abs() < (ba - ih / iw.max(1.0)).abs()
    };
    match turn_deg {
        90 if !box_unturned => img.rotate90(),
        180 => img.rotate180(),
        270 if !box_unturned => img.rotate270(),
        _ => img,
    }
}

/// `font::runtime::face_file`, remembered for the process. The lookup reads every
/// font file on the machine when the face is absent, and the text arm asks once per
/// bold East Asian element: blind-J ja/educational__28102e23 spent 9 of its 10 minutes
/// there (layout 6s). Same answers, asked once per (family, bold, italic).
fn face_file_cached(family: &str, bold: bool, italic: bool) -> Option<(std::path::PathBuf, u32)> {
    use std::sync::{Mutex, OnceLock};
    static CACHE: OnceLock<Mutex<HashMap<(String, bool, bool), Option<(std::path::PathBuf, u32)>>>> = OnceLock::new();
    let cache = CACHE.get_or_init(|| Mutex::new(HashMap::new()));
    let key = (family.to_string(), bold, italic);
    if let Some(hit) = cache.lock().unwrap().get(&key) {
        return hit.clone();
    }
    let found = oxidocs_core::font::runtime::face_file(family, bold, italic);
    cache.lock().unwrap().insert(key, found.clone());
    found
}

fn cjk_bold_face_enabled() -> bool {
    std::env::var_os("OXI_PDF_CJK_BOLD_FACE_DISABLE").is_none()
}

fn resolve_font_styled(family: Option<&str>, bold: bool, italic: bool) -> String {
    let base = family.unwrap_or("Calibri");
    // CJK fonts: use bold variant when bold is requested
    if is_cjk_family(base) {
        if bold {
            // OXI_PDF_CJK_BOLD_FACE (2026-11-09, default ON, opt-out _DISABLE): a family
            // with its own bold face on this machine (メイリオ → meiryob.ttc, 游ゴシック
            // Bold, BIZ UD…) is drawn with it, embedded as «<name>-Bold»; only a family
            // without one goes to the shared stand-in. blind-J ja/creative__99e3d9f9: bold
            // メイリオ «2024 年 10 月 24 日» came out in ＭＳ ゴシック (Word: Meiryo-Bold).
            if cjk_bold_face_enabled() && face_file_cached(base, true, false).is_some() {
                return format!("{}-Bold", resolve_font_styled(Some(base), false, italic));
            }
            return "OxiCJK-Bold".to_string();
        }
        return match base {
            "ＭＳ ゴシック" | "MS ゴシック" | "MS Gothic" => "MS-Gothic".to_string(),
            "ＭＳ Ｐゴシック" | "MS PGothic" => "MS-PGothic".to_string(),
            "ＭＳ 明朝" | "MS 明朝" | "MS Mincho" => "MS-Mincho".to_string(),
            "ＭＳ Ｐ明朝" | "MS PMincho" => "MS-PMincho".to_string(),
            _ => base.to_string(),
        };
    }
    // A face the plan actually embedded wins over the base-14 mapping: the
    // document's own family drawn with its own outlines and widths, rather
    // than every sans family collapsed onto one file.
    if let Some(key) = FONT_PLAN.with(|plan| plan.borrow().keys.get(&(base.trim().to_lowercase(), bold, italic)).cloned()) {
        return key;
    }

    // Nothing embedded for this face: draw it with an embedded stand-in. Times
    // goes there too, not to base-14 Times-Roman: its substitute is only
    // missing here when the run asks for glyphs it lacks, and a base-14 name
    // is a font with no file, which a PDF/A check rejects.
    if bold {
        match base {
            "Calibri" | "Arial" | "Helvetica" => "OxiLatin-Bold".to_string(),
            "Cambria" | "Century" => "OxiCambria-Bold".to_string(),
            "Courier New" | "Courier" => "Courier-Bold".to_string(),
            other => stand_in_face(other, true),
        }
    } else {
        match base {
            "Calibri" | "Arial" | "Helvetica" => "OxiLatin-Regular".to_string(),
            "Cambria" | "Century" => "OxiCambria-Regular".to_string(),
            "Courier New" => "Courier".to_string(),
            other => stand_in_face(other, false),
        }
    }
}

/// The always-embedded face that draws a family this machine does not have
/// and no base-14 name stands for: Verdana, Aptos, Segoe UI, Wingdings and
/// the like. The family's own name used to go into the PDF instead, as a
/// simple font with no file and no width table. A viewer then drew it with
/// whatever it picked, at that face's widths; a PDF/A check rejects it
/// outright; and a symbol font's codes, mapped to Unicode for want of the
/// face, came out as whatever those bytes mean in WinAnsi. An embedded
/// stand-in draws the same glyphs everywhere, and what it lacks goes on to
/// the fallback faces. Serif families get the serif stand-in, the rest the
/// sans one.
fn stand_in_face(family: &str, bold: bool) -> String {
    let lower = family.to_lowercase();
    const SERIF: &[&str] = &[
        "serif", "times", "roman", "georgia", "garamond", "book", "century", "cambria", "palatino",
        "baskerville", "bodoni", "caslon", "minion", "constantia", "didot", "perpetua", "rockwell",
        "sabon", "goudy", "cochin", "charter", "mincho", "ming", "song",
    ];
    // Century Gothic and Franklin Gothic are sans faces whatever else their
    // names hold.
    let serif = SERIF.iter().any(|k| lower.contains(k)) && !lower.contains("sans") && !lower.contains("gothic");
    match (serif, bold) {
        (true, true) => "OxiCambria-Bold",
        (true, false) => "OxiCambria-Regular",
        (false, true) => "OxiLatin-Bold",
        (false, false) => "OxiLatin-Regular",
    }
    .to_string()
}

fn parse_hex_color(hex: &str) -> Option<Color> {
    let hex = hex.strip_prefix('#').unwrap_or(hex);
    if hex.len() != 6 { return None; }
    let r = u8::from_str_radix(&hex[0..2], 16).ok()? as f64 / 255.0;
    let g = u8::from_str_radix(&hex[2..4], 16).ok()? as f64 / 255.0;
    let b = u8::from_str_radix(&hex[4..6], 16).ok()? as f64 / 255.0;
    Some(Color::Rgb(r, g, b))
}

fn parse_highlight_color(name: &str) -> Option<Color> {
    match name {
        "yellow" => Some(Color::Rgb(1.0, 1.0, 0.0)),
        "green" => Some(Color::Rgb(0.0, 1.0, 0.0)),
        "cyan" => Some(Color::Rgb(0.0, 1.0, 1.0)),
        "magenta" => Some(Color::Rgb(1.0, 0.0, 1.0)),
        "blue" => Some(Color::Rgb(0.0, 0.0, 1.0)),
        "red" => Some(Color::Rgb(1.0, 0.0, 0.0)),
        "darkBlue" => Some(Color::Rgb(0.0, 0.0, 0.55)),
        "darkCyan" => Some(Color::Rgb(0.0, 0.55, 0.55)),
        "darkGreen" => Some(Color::Rgb(0.0, 0.39, 0.0)),
        "darkMagenta" => Some(Color::Rgb(0.55, 0.0, 0.55)),
        "darkRed" => Some(Color::Rgb(0.55, 0.0, 0.0)),
        "darkYellow" => Some(Color::Rgb(0.55, 0.55, 0.0)),
        "darkGray" | "darkGrey" => Some(Color::Rgb(0.66, 0.66, 0.66)),
        "lightGray" | "lightGrey" => Some(Color::Rgb(0.85, 0.85, 0.85)),
        "black" => Some(Color::Gray(0.0)),
        _ => parse_hex_color(name),
    }
}

/// Detect paragraph "role" from text patterns and run styling
#[derive(Debug, Clone, Copy, PartialEq)]
enum ParaRole {
    CoverTitle,      // First centered paragraph with large font
    CoverSubtitle,   // Second centered paragraph
    CoverAuthor,     // Third centered paragraph (author/date)
    SectionHeader,   // ① ② ③ ... ⑩ section headers
    SubHeader,       // "N-N." sub-headers
    SubSubHeader,    // Bold paragraph with spB>=200 and smaller font (topic intro)
    Body,            // Regular body text
    ListItem,        // Bulleted/numbered list
    CreditFooter,    // Last right-aligned "本提案書の..." line
}

fn classify_paragraph(para: &Paragraph, block_index: usize, _total_blocks: usize) -> ParaRole {
    let text: String = para.runs.iter().map(|r| r.text.as_str()).collect();
    let trimmed = text.trim();

    // Credit footer (last non-empty paragraph, right-aligned)
    if para.alignment == Alignment::Right && trimmed.contains("本提案書") {
        return ParaRole::CreditFooter;
    }

    // Cover page: first 3 centered paragraphs
    if para.alignment == Alignment::Center && block_index <= 2 {
        let font_size = para.runs.first()
            .and_then(|r| r.style.font_size)
            .unwrap_or(10.5);
        if block_index == 0 && font_size >= 14.0 {
            return ParaRole::CoverTitle;
        }
        if block_index == 1 {
            return ParaRole::CoverSubtitle;
        }
        if block_index == 2 {
            return ParaRole::CoverAuthor;
        }
    }

    // Section headers: starts with ① ② ③ etc.
    if trimmed.starts_with('①') || trimmed.starts_with('②') || trimmed.starts_with('③')
        || trimmed.starts_with('④') || trimmed.starts_with('⑤') || trimmed.starts_with('⑥')
        || trimmed.starts_with('⑦') || trimmed.starts_with('⑧') || trimmed.starts_with('⑨')
        || trimmed.starts_with('⑩')
    {
        return ParaRole::SectionHeader;
    }

    // Sub-headers: patterns like "1-1.", "2-2.", "3-1." etc.
    let is_sub_header = {
        let chars: Vec<char> = trimmed.chars().collect();
        chars.len() >= 4
            && chars[0].is_ascii_digit()
            && chars[1] == '-'
            && chars[2].is_ascii_digit()
            && chars[3] == '.'
    };
    if is_sub_header {
        return ParaRole::SubHeader;
    }

    // Sub-sub-headers: bold text with space_before >= 150, font size ~11pt, not too long
    let font_size = para.runs.first()
        .and_then(|r| r.style.font_size)
        .or(para.style.default_run_style.as_ref().and_then(|rs| rs.font_size))
        .unwrap_or(10.5);
    let is_bold = para.runs.first().map(|r| r.style.bold).unwrap_or(false);
    let space_before = para.style.space_before.unwrap_or(0.0);

    if is_bold && space_before >= 10.0 && font_size >= 10.5 && trimmed.len() < 120
        && !trimmed.starts_with('・') && !trimmed.starts_with('【')
        && para.alignment != Alignment::Center
        && para.style.list_marker.is_none()
    {
        return ParaRole::SubSubHeader;
    }

    // List items
    if para.style.list_marker.is_some() || trimmed.starts_with('・') || trimmed.starts_with('※') {
        return ParaRole::ListItem;
    }

    ParaRole::Body
}

/// The PDF, and the header and footer parts its pages draw.
fn doc_to_pdf(doc: &oxidocs_core::Document) -> (PdfDocument, std::collections::BTreeSet<String>) {
    // Use LayoutEngine for layout (single source of truth with WASM path)
    let engine = oxidocs_core::layout::LayoutEngine::for_document(doc);
    let layout_result = engine.layout(doc);

    // Font metrics for accurate baseline calculation
    let registry = FontMetricsRegistry::load();

    // Build embedded fonts for CJK text
    let mut embedded_fonts = build_embedded_fonts(doc, &layout_result);
    embed_fallback_faces(&layout_result, &mut embedded_fonts);

    // Convert LayoutResult pages → PDF pages
    let pages: Vec<Page> = layout_result.pages.iter().map(|lp| {
        let w = lp.width as f64;
        let h = lp.height as f64;
        let mut contents: Vec<ContentElement> = Vec::new();
        // (index in contents, left, top, bottom, right, highlight name) of the last highlight box
        let mut last_hl: Option<(usize, f64, f64, f64, f64, String)> = None;

        // OXI_PDF_SYMBOL_MARKER_BASELINE (2026-11-13, default ON, opt-out _DISABLE): a Symbol-font
        // list marker sits on its line's body baseline, as the DirectWrite renderer already draws it
        // (S518). The PDF placed it on Symbol's own ascent: blind-J en/policies__00d0246e (8pt
        // bullet, 11pt Calibri Light) 2.43pt high, educational__007b4139 / policies__00c94b4f (same
        // size) 0.5pt; Word's PDF puts every bullet origin on the text's (13 docs, 0.00).
        let marker_body_ref: std::collections::HashMap<i32, (f32, Option<String>, f32)> =
            if std::env::var_os("OXI_PDF_SYMBOL_MARKER_BASELINE_DISABLE").is_none() {
                let mut m = std::collections::HashMap::new();
                for e in &lp.elements {
                    if e.paragraph_index.is_none() { continue; }
                    if let oxidocs_core::layout::LayoutContent::Text { text, font_size, font_family, .. } = &e.content {
                        if text.trim().is_empty() { continue; }
                        if font_family.as_deref().map(oxidocs_core::font::render_family_name) == Some("Symbol") { continue; }
                        m.entry((e.y * 4.0).round() as i32)
                            .or_insert((e.text_y_off, font_family.clone(), *font_size));
                    }
                }
                m
            } else { std::collections::HashMap::new() };

        for elem in &lp.elements {
            let x = elem.x as f64;
            let y = elem.y as f64;
            let ew = elem.width as f64;
            let eh = elem.height as f64;

            match &elem.content {
                oxidocs_core::layout::LayoutContent::Text {
                    text, font_size, font_family, bold, italic,
                    underline, underline_style, strikethrough, color, highlight,
                    character_spacing, text_scale, effects, is_vertical, ..
                } => {
                    // OXI_PDF_TEXT_EFFECTS (2026-11-08, default ON, opt-out _DISABLE): Word's
                    // character effects as the DirectWrite renderer draws them -- no fill
                    // paints nothing; outline = a white fill ringed in the text colour;
                    // shadow = a gray copy 1pt down-right underneath; a w14:textOutline is
                    // the ring in its own colour. The CLI drew the plain colour (blind-J
                    // ja/creative__9aa11078: the white outlined title came out blue on blue).
                    let fx_on = std::env::var_os("OXI_PDF_TEXT_EFFECTS_DISABLE").is_none();
                    if fx_on && effects.no_fill {
                        continue;
                    }
                    let fs = *font_size as f64;
                    // OXI_PDF_TEXT_SCALE (2026-11-08, default ON, opt-out _DISABLE): w:w is
                    // drawn as a horizontal glyph scale (Word's PDF: a compressed text
                    // matrix); without it the glyphs kept full width at the layout's narrower
                    // positions and overlapped (blind-J ja/creative__9aa11078, w:w=66).
                    let hscale = if std::env::var_os("OXI_PDF_TEXT_SCALE_DISABLE").is_none()
                        && *text_scale > 0.0 { *text_scale as f64 } else { 100.0 };
                    let fill_color = color.as_deref()
                        .and_then(parse_hex_color)
                        .unwrap_or(Color::Gray(0.0));

                    // Highlight background
                    if let Some(ref hl) = highlight {
                        // OXI_PDF_HIGHLIGHT_JOIN (2026-11-09, default ON, opt-out _DISABLE): a
                        // highlight that continues the previous one on the same line widens
                        // that box instead of starting another; abutting boxes left a light
                        // anti-aliased seam at every element edge (blind-J ja/policies__28f59048,
                        // «中央太郎（チュウオウタロウ）» striped where Word draws one band).
                        let join = std::env::var_os("OXI_PDF_HIGHLIGHT_JOIN_DISABLE").is_none()
                            && last_hl.as_ref().is_some_and(|(_, _, t, b, r, name)| name == hl
                                && (t - y).abs() < 0.01 && (b - (y + eh)).abs() < 0.01
                                && x >= *r - 0.75 && x <= *r + 0.75);
                        if join {
                            let (idx, l, t, b, r, _) = last_hl.as_mut().unwrap();
                            *r = r.max(x + ew);
                            if let ContentElement::Path(p) = &mut contents[*idx] {
                                p.operations = vec![
                                    PathOp::MoveTo(*l, *t),
                                    PathOp::LineTo(*r, *t),
                                    PathOp::LineTo(*r, *b),
                                    PathOp::LineTo(*l, *b),
                                    PathOp::ClosePath,
                                ];
                            }
                        } else if let Some(hl_color) = parse_highlight_color(hl) {
                            last_hl = Some((contents.len(), x, y, y + eh, x + ew, hl.clone()));
                            contents.push(ContentElement::Path(PathData {
                                operations: vec![
                                    PathOp::MoveTo(x, y),
                                    PathOp::LineTo(x + ew, y),
                                    PathOp::LineTo(x + ew, y + eh),
                                    PathOp::LineTo(x, y + eh),
                                    PathOp::ClosePath,
                                ],
                                stroke: None,
                                fill: Some(hl_color),
                            }));
                        }
                    }

                    // OXI_PDF_FAKE_BOLD (2026-11-08, default ON, opt-out _DISABLE): an East
                    // Asian family with no bold member on this machine is drawn as Word draws
                    // it -- its regular face, filled and stroked at fs/35 (Word's PDF: «2 Tr
                    // 0.51429 w» at 18pt) -- instead of a stand-in face. blind-J ja/reports__
                    // 77526bdb's bold ＭＳ Ｐゴシック «長野労働局» came out light.
                    let fake_bold = *bold
                        && std::env::var_os("OXI_PDF_FAKE_BOLD_DISABLE").is_none()
                        && font_family.as_deref().is_some_and(|f| is_cjk_family(f)
                            && face_file_cached(f, true, false).is_none());
                    // Resolve font name for PDF (map to Type1 base fonts or CJK)
                    let font_name = resolve_font_styled(font_family.as_deref(), *bold && !fake_bold, *italic);
                    let font_name = if *bold && !fake_bold && font_family.as_deref().is_some_and(is_cjk_family)
                        && font_name != "OxiCJK-Bold" && !embedded_fonts.contains_key(&font_name)
                    {
                        "OxiCJK-Bold".to_string()
                    } else {
                        font_name
                    };
                    let text = &symbol_text(font_family.as_deref(), &font_name, text, &embedded_fonts);
                    let chars: Vec<char> = text.chars().collect();
                    let mut seg_x = x;
                    let mut segments: Vec<(String, String, f64, f64)> = Vec::new();

                    // Consecutive characters drawn with the same face make one
                    // segment: the run's own face, the East Asian face for what
                    // a Latin face lacks, a fallback face for what both lack.
                    let mut raw_segments: Vec<(String, String, f64)> = Vec::new();
                    let mut total_est = 0.0_f64;
                    let mut i = 0;
                    while i < chars.len() {
                        let seg_font = pdf_face_for(chars[i], &font_name, *bold, &embedded_fonts);
                        let mut j = i + 1;
                        while j < chars.len() && pdf_face_for(chars[j], &font_name, *bold, &embedded_fonts) == seg_font {
                            j += 1;
                        }
                        let seg_text: String = chars[i..j].iter().collect();
                        let seg_w: f64 = seg_text.chars()
                            .map(|c| char_width(c, fs))
                            .sum();
                        total_est += seg_w;
                        raw_segments.push((seg_text, seg_font, seg_w));
                        i = j;
                    }

                    // Second pass: scale segment widths to match LayoutEngine's elem.width
                    let scale = if total_est > 0.0 { ew / total_est } else { 1.0 };
                    for (seg_text, seg_font, seg_w) in raw_segments {
                        let scaled_w = seg_w * scale;
                        segments.push((seg_text, seg_font, seg_x, scaled_w));
                        seg_x += scaled_w;
                    }

                    // Baseline position: y is top of line box, PDF baseline = y + ascent
                    // Matches Word output: baseline = content_top + word_ascent_pt(fontSize)
                    let text_y = {
                        let metrics = font_family.as_deref()
                            .map(|ff| registry.get(ff))
                            .unwrap_or_else(|| registry.default_metrics());
                        // OXI_PDF_BASELINE (2026-11-08, default ON, opt-out _DISABLE): the
                        // PDF puts the baseline where the DirectWrite renderer does (the
                        // renderer the README scores): the layout's own baseline when it
                        // carries one, else the glyph top
                        // (line top + text_y_off − 1pt, the renderer's origin) plus the metric table's
                        // baseline ascent.
                        // blind-J ja (dev pool, Oxi 7774ac9c, Windows): this CLI's PDF sat
                        // 0.1em low on MS Mincho (dy +1.14pt at 10.5, +0.91 at 9, +0.79 at
                        // 8) and 0.6pt high on Calibri against Word's PDF; penalized SSIM
                        // JA 0.8364 vs 0.9282 through the renderer, EN 0.9466 vs 0.9565.
                        let symbol_marker_body = if elem.paragraph_index.is_none()
                            && elem.baseline_offset.is_none()
                            && font_family.as_deref().map(oxidocs_core::font::render_family_name) == Some("Symbol")
                        {
                            marker_body_ref.get(&((elem.y * 4.0).round() as i32))
                        } else { None };
                        if let (true, Some((body_off, body_fam, body_fs))) =
                            (std::env::var_os("OXI_PDF_BASELINE_DISABLE").is_none(), symbol_marker_body)
                        {
                            let bm = body_fam.as_deref()
                                .map(|ff| registry.get(ff))
                                .unwrap_or_else(|| registry.default_metrics());
                            y + *body_off as f64 - 1.0 + (bm.baseline_ascent() * *body_fs) as f64
                        } else if std::env::var_os("OXI_PDF_BASELINE_DISABLE").is_none() {
                            match elem.baseline_offset {
                                Some(b) => y + b as f64,
                                None => y + elem.text_y_off as f64 - 1.0
                                    + (metrics.baseline_ascent() * fs as f32) as f64,
                            }
                        } else {
                            y + metrics.word_ascent_pt(fs as f32) as f64
                        }
                    };

                    // OXI_PDF_VERT_GLYPHS (2026-11-09, default ON, opt-out _DISABLE): vertical
                    // text is drawn as the DirectWrite renderer draws it -- each character
                    // upright, one em (+ spacing) further down, centred in the element's
                    // column, with the face's own `vert` glyph (、。 in the upper right,
                    // brackets and ー turned). The CLI drew the horizontal forms
                    // (blind-J ja/creative__95f17b17, a tbRl novel).
                    if *is_vertical && std::env::var_os("OXI_PDF_VERT_GLYPHS_DISABLE").is_none() {
                        let adv = fs + *character_spacing as f64;
                        let cx = x + ((ew - fs) * 0.5).max(0.0);
                        for (k, ch) in chars.iter().enumerate() {
                            if ch.is_whitespace() {
                                continue;
                            }
                            let key = pdf_face_for(*ch, &font_name, *bold, &embedded_fonts);
                            let vg = embedded_fonts.get(&key)
                                .and_then(|f| f.unicode_to_gid.get(&(VERT_GLYPH_KEY | *ch as u32)).copied());
                            let span = TextSpan {
                                x: cx, y: text_y + k as f64 * adv,
                                text: ch.to_string(), font_name: key,
                                font_size: fs, fill_color,
                                character_spacing: 0.0, stroke_width: if fake_bold { fs / 35.0 } else { 0.0 }, skew: 0.0, horizontal_scale: 100.0, stroke_color: None,
                            };
                            contents.push(match vg {
                                Some(g) => ContentElement::GlyphRun(GlyphRun { span, glyph_indices: vec![g] }),
                                None => ContentElement::Text(span),
                            });
                        }
                        continue;
                    }
                    let hex = |c: u32| Color::Rgb(((c >> 16) & 255) as f64 / 255.0, ((c >> 8) & 255) as f64 / 255.0, (c & 255) as f64 / 255.0);
                    // (fill, stroke width, stroke colour) of the main pass
                    // OXI_PDF_W14_OUTLINE_WIDTH (2026-11-13, default ON, opt-out _DISABLE): a run with
                    // both w:outline and a w14:textOutline is ringed at the w14 width and colour, as
                    // Word's PDF draws it («w0.8 RG(4472C4)» for textOutline w=10160 at 48 and 18pt);
                    // the fixed 1.6 swallowed the small ruby's dakuten (blind-J ja/creative__9aa11078
                    // «こうざ» read «こうさ»).
                    let (main_fill, main_w, main_stroke) = if fx_on && effects.outline
                        && std::env::var_os("OXI_PDF_W14_OUTLINE_WIDTH_DISABLE").is_none()
                        && effects.line.is_some()
                    {
                        let (lw, lc) = effects.line.unwrap();
                        (Color::Gray(1.0), (lw as f64).max(0.25), Some(hex(lc)))
                    } else if fx_on && effects.outline {
                        (Color::Gray(1.0), 1.6, Some(fill_color))
                    } else if let (true, Some((lw, lc))) = (fx_on, effects.line) {
                        (fill_color, (lw as f64).max(0.5), Some(hex(lc)))
                    } else if fake_bold {
                        (fill_color, fs / 35.0, None)
                    } else {
                        (fill_color, 0.0, None)
                    };
                    // OXI_PDF_BLANK_SEGMENT (2026-11-09, default ON, opt-out _DISABLE): a
                    // segment of only tabs or breaks draws no glyph -- blank as in Word; the
                    // face has none for it and the PDF showed .notdef boxes (blind-J ja/
                    // creative__9aa11078: a tab element drew a «?» box at the top left of
                    // page 1). Its width still carries the underline.
                    // OXI_PDF_SYNTH_ITALIC (2026-11-09, default ON, opt-out _DISABLE): italic
                    // text in an East Asian family (drawn with its regular face) is sheared
                    // as Word draws it -- Tm «1 0 0.3333 1» in Word's PDF. zh dev
                    // legal__00849104's bold-italic 宋体 headings «监管协议» came out upright.
                    let synth_skew = if *italic
                        && std::env::var_os("OXI_PDF_SYNTH_ITALIC_DISABLE").is_none()
                        && font_family.as_deref().is_some_and(is_cjk_family)
                    { 0.3333 } else { 0.0 };
                    let blank_skip = std::env::var_os("OXI_PDF_BLANK_SEGMENT_DISABLE").is_none();
                    for (seg_text, seg_font, sx, _sw) in segments {
                        if blank_skip && seg_text.chars().all(|c| c.is_control()) {
                            continue;
                        }
                        // OXI_PDF_PUA_BLANK (2026-11-10, default ON, opt-out _DISABLE): a
                        // private-use character that no face in the PDF covers is blank, as Word
                        // draws it (its fallback face shows nothing); the stand-in face's .notdef
                        // drew a box. zh dev creative__0f11c107: every bullet item opens with
                        // U+F06C in DengXian -- Word's PDF shows it in Calibri with no ink, the
                        // CLI drew «□» after each «●».
                        if std::env::var_os("OXI_PDF_PUA_BLANK_DISABLE").is_none()
                            && seg_text.chars().all(|c| ('\u{E000}'..='\u{F8FF}').contains(&c))
                            && embedded_fonts.get(&seg_font).map_or(true, |f| {
                                seg_text.chars().all(|c| !f.unicode_to_gid.contains_key(&(c as u32)))
                            })
                        {
                            continue;
                        }
                        // OXI_PDF_SYNTH_ITALIC_LATIN (2026-11-13, default ON, opt-out _DISABLE):
                        // italic text whose drawn face has no italic of its own is sheared
                        // 1/3 like the East Asian case -- Word's PDF writes the same
                        // «1 0 0.3333 1» Tm. blind-J en/policies__00cb6db7: «Letter-join 16»
                        // (altName Eras Demi ITC, no italic face) italic paragraphs came out
                        // upright; Word's PDF has 29 sheared text matrices on that page.
                        let synth_skew = if synth_skew == 0.0
                            && *italic
                            && std::env::var_os("OXI_PDF_SYNTH_ITALIC_LATIN_DISABLE").is_none()
                            && embedded_fonts.get(&seg_font)
                                .and_then(|f| rustybuzz::Face::from_slice(&f.data, 0))
                                .is_some_and(|f| !f.is_italic() && !f.is_oblique())
                        { 0.3333 } else { synth_skew };
                        if fx_on && effects.shadow {
                            contents.push(ContentElement::Text(TextSpan {
                                x: sx + 1.0, y: text_y + 1.0,
                                text: seg_text.clone(), font_name: seg_font.clone(),
                                font_size: fs, fill_color: Color::Gray(0.5),
                                character_spacing: *character_spacing as f64, stroke_width: 0.0, skew: synth_skew, horizontal_scale: hscale, stroke_color: None,
                            }));
                        }
                        contents.push(ContentElement::Text(TextSpan {
                            x: sx, y: text_y,
                            text: seg_text, font_name: seg_font,
                            font_size: fs, fill_color: main_fill,
                            character_spacing: *character_spacing as f64, stroke_width: main_w, skew: synth_skew, horizontal_scale: hscale, stroke_color: main_stroke,
                        }));
                    }

                    // Underline
                    if *underline {
                        let is_double = underline_style.as_deref() == Some("double");
                        let ul_w = 0.5_f64.max(fs * 0.05);
                        if is_double {
                            // Double underline: two lines
                            let ul_y1 = text_y + fs * 0.10;
                            let ul_y2 = text_y + fs * 0.22;
                            for ul_y in [ul_y1, ul_y2] {
                                contents.push(ContentElement::Path(PathData {
                                    operations: vec![
                                        PathOp::MoveTo(x, ul_y),
                                        PathOp::LineTo(seg_x, ul_y),
                                    ],
                                    stroke: Some(StrokeStyle {
                                        color: fill_color, width: ul_w,
                                        line_cap: LineCap::Butt, line_join: LineJoin::Miter,
                                    }),
                                    fill: None,
                                }));
                            }
                        } else {
                            let ul_y = text_y + fs * 0.15;
                            contents.push(ContentElement::Path(PathData {
                                operations: vec![
                                    PathOp::MoveTo(x, ul_y),
                                    PathOp::LineTo(seg_x, ul_y),
                                ],
                                stroke: Some(StrokeStyle {
                                    color: fill_color, width: ul_w,
                                    line_cap: LineCap::Butt, line_join: LineJoin::Miter,
                                }),
                                fill: None,
                            }));
                        }
                    }

                    // Strikethrough
                    if *strikethrough {
                        let st_y = y + eh * 0.45;
                        let st_w = 0.5_f64.max(fs * 0.04);
                        contents.push(ContentElement::Path(PathData {
                            operations: vec![
                                PathOp::MoveTo(x, st_y),
                                PathOp::LineTo(seg_x, st_y),
                            ],
                            stroke: Some(StrokeStyle {
                                color: fill_color, width: st_w,
                                line_cap: LineCap::Butt, line_join: LineJoin::Miter,
                            }),
                            fill: None,
                        }));
                    }
                }

                oxidocs_core::layout::LayoutContent::TableBorder { x1, y1, x2, y2, color, width, .. } => {
                    let bc = color.as_deref()
                        .and_then(parse_hex_color)
                        .unwrap_or(Color::Gray(0.0));
                    let bw = *width as f64;
                    // OXI_PDF_HRULE_BELOW (2026-11-09, default ON, opt-out _DISABLE): a
                    // horizontal rule hangs below its coordinate -- Word's PDF fills
                    // [y, y + w] (blind-J page 1, ~25 docs EN+JA: table rules' centres
                    // sit +0.5w below the layout's y; policies__28f59048 28.44 / 67.32 /
                    // 417.19 for y 28.40 / 67.35 / 417.15). Vertical rules stay centred.
                    let dy = if (*y1 - *y2).abs() < 0.01
                        && std::env::var_os("OXI_PDF_HRULE_BELOW_DISABLE").is_none() { bw * 0.5 } else { 0.0 };
                    contents.push(ContentElement::Path(PathData {
                        operations: vec![
                            PathOp::MoveTo(*x1 as f64, *y1 as f64 + dy),
                            PathOp::LineTo(*x2 as f64, *y2 as f64 + dy),
                        ],
                        stroke: Some(StrokeStyle {
                            color: bc, width: bw,
                            line_cap: LineCap::Butt, line_join: LineJoin::Miter,
                        }),
                        fill: None,
                    }));
                }

                oxidocs_core::layout::LayoutContent::CellShading { color: shade_color } => {
                    let sc = parse_hex_color(shade_color).unwrap_or(Color::Gray(0.9));
                    contents.push(ContentElement::Path(PathData {
                        operations: vec![
                            PathOp::MoveTo(x, y),
                            PathOp::LineTo(x + ew, y),
                            PathOp::LineTo(x + ew, y + eh),
                            PathOp::LineTo(x, y + eh),
                            PathOp::ClosePath,
                        ],
                        stroke: None,
                        fill: Some(sc),
                    }));
                }

                oxidocs_core::layout::LayoutContent::BoxRect { fill, stroke_color, stroke_width, corner_radius } => {
                    let fill_color = fill.as_deref().and_then(parse_hex_color);
                    let stroke = stroke_color.as_deref().and_then(parse_hex_color)
                        .map(|c| (*stroke_width as f64, c));
                    let cr = *corner_radius as f64;
                    if cr > 0.0 {
                        // Rounded rectangle
                        let k = 0.5523 * cr;
                        let (rx, ry, rw, rh) = (x, y, ew, eh);
                        let ops = vec![
                            PathOp::MoveTo(rx + cr, ry),
                            PathOp::LineTo(rx + rw - cr, ry),
                            PathOp::CurveTo(rx + rw - cr + k, ry, rx + rw, ry + cr - k, rx + rw, ry + cr),
                            PathOp::LineTo(rx + rw, ry + rh - cr),
                            PathOp::CurveTo(rx + rw, ry + rh - cr + k, rx + rw - cr + k, ry + rh, rx + rw - cr, ry + rh),
                            PathOp::LineTo(rx + cr, ry + rh),
                            PathOp::CurveTo(rx + cr - k, ry + rh, rx, ry + rh - cr + k, rx, ry + rh - cr),
                            PathOp::LineTo(rx, ry + cr),
                            PathOp::CurveTo(rx, ry + cr - k, rx + cr - k, ry, rx + cr, ry),
                            PathOp::ClosePath,
                        ];
                        contents.push(ContentElement::Path(PathData {
                            operations: ops,
                            stroke: stroke.map(|(sw, sc)| StrokeStyle {
                                color: sc, width: sw, line_cap: LineCap::Butt, line_join: LineJoin::Round,
                            }),
                            fill: fill_color,
                        }));
                    } else {
                        // Regular rectangle
                        let ops = vec![
                            PathOp::MoveTo(x, y),
                            PathOp::LineTo(x + ew, y),
                            PathOp::LineTo(x + ew, y + eh),
                            PathOp::LineTo(x, y + eh),
                            PathOp::ClosePath,
                        ];
                        contents.push(ContentElement::Path(PathData {
                            operations: ops,
                            stroke: stroke.map(|(sw, sc)| StrokeStyle {
                                color: sc, width: sw, line_cap: LineCap::Butt, line_join: LineJoin::Miter,
                            }),
                            fill: fill_color,
                        }));
                    }
                }

                oxidocs_core::layout::LayoutContent::Image { ref data, ref content_type, crop, rotation, .. } => {
                    // Decode image (PNG/JPEG/etc.) to raw RGB pixels
                    if !data.is_empty() {
                        // OXI_PDF_EMF: a metafile is played as vectors (emf.rs).
                        if emf::is_emf(data) || emf::is_wmf(data) {
                            let key = |face: &str, bold: bool, italic: bool, c: char| {
                                pdf_face_for(c, &resolve_font_styled(Some(face), bold, italic), bold, &embedded_fonts)
                            };
                            let metrics = |face: &str, _bold: bool| {
                                let m = registry.get(face);
                                (m.win_ascent as f64, m.win_descent as f64)
                            };
                            let faces = emf::TextFaces { key: &key, metrics: &metrics };
                            let (ex, ey, ew2, eh2) = (elem.x as f64, elem.y as f64, elem.width as f64, elem.height as f64);
                            let played = if emf::is_emf(data) {
                                emf::render(data, ex, ey, ew2, eh2, Some(&faces), *crop)
                            } else {
                                emf::render_wmf(data, ex, ey, ew2, eh2, Some(&faces), *crop)
                            };
                            if let Some(els) = played {
                                contents.extend(els);
                            }
                        } else if let Ok(img) = image::load_from_memory(data) {
                            // OXI_PDF_IMAGE_CROP (2026-11-09, default ON, opt-out _DISABLE):
                            // a:srcRect crops the source and quarter turns turn the pixels,
                            // as the DirectWrite renderer does (S775 / OXI_CROP_NEGATIVE /
                            // OXI_PICTURE_ROTATION). The PDF drew the whole picture squeezed
                            // into the box: blind-J en/policies__00d4d26d's footer skyline
                            // (srcRect t 57.611 r -0.496) on every page.
                            let img = if std::env::var_os("OXI_PDF_IMAGE_CROP_DISABLE").is_none() {
                                pdf_image_crop_turn(img, *crop, *rotation, elem.width, elem.height)
                            } else {
                                img
                            };
                            // OXI_PDF_IMAGE_ALPHA (2026-11-08, default ON, opt-out _DISABLE):
                            // an image with transparent pixels keeps them as a soft mask;
                            // to_rgb8 alone turned them black (blind-J ja/reports__77526bdb's
                            // transparent PNG illustration drew on a black square).
                            let alpha = if img.color().has_alpha()
                                && std::env::var_os("OXI_PDF_IMAGE_ALPHA_DISABLE").is_none()
                            {
                                let a: Vec<u8> = img.to_rgba8().pixels().map(|p| p.0[3]).collect();
                                a.iter().any(|&v| v < 255).then_some(a)
                            } else {
                                None
                            };
                            let rgb = img.to_rgb8();
                            let (pw, ph) = rgb.dimensions();
                            let raw_pixels = rgb.into_raw();
                            contents.push(ContentElement::Image(ImageData {
                                x: elem.x as f64,
                                y: elem.y as f64,
                                width: elem.width as f64,
                                height: elem.height as f64,
                                data: raw_pixels,
                                color_space: ColorSpace::DeviceRgb,
                                bits_per_component: 8,
                                pixel_width: pw,
                                pixel_height: ph,
                                alpha,
                            }));
                        }
                    }
                }
                oxidocs_core::layout::LayoutContent::ClipStart => {
                    // Pass coordinates in page-top-down system; the PDF writer's
                    // ClipPath handler flips Y via (page.height - y).
                    let clip_x = elem.x as f64;
                    let clip_y = elem.y as f64;
                    let clip_w = elem.width as f64;
                    let clip_h = elem.height as f64;
                    contents.push(ContentElement::SaveState);
                    contents.push(ContentElement::ClipPath(ClipPathData {
                        operations: vec![
                            PathOp::MoveTo(clip_x, clip_y),
                            PathOp::LineTo(clip_x + clip_w, clip_y),
                            PathOp::LineTo(clip_x + clip_w, clip_y + clip_h),
                            PathOp::LineTo(clip_x, clip_y + clip_h),
                            PathOp::ClosePath,
                        ],
                        even_odd: false,
                    }));
                }
                oxidocs_core::layout::LayoutContent::ClipEnd => {
                    contents.push(ContentElement::RestoreState);
                }
                oxidocs_core::layout::LayoutContent::WatermarkText { .. } => {}
                oxidocs_core::layout::LayoutContent::PresetShape { ref shape_type, ref stroke_color, stroke_width, flip_h, flip_v, arrow_head, arrow_tail, .. }
                    if std::env::var_os("OXI_PDF_PRESET_DISABLE").is_none() =>
                {
                    // OXI_PDF_PRESET (2026-11-08, default ON, opt-out _DISABLE): draw a
                    // preset shape as the DirectWrite renderer does -- connectors as
                    // flip-aware lines with filled triangle ends, ellipses as rings,
                    // brackets as brackets, and nothing for any other preset (its body
                    // is drawn by its own box / path element). Every other preset came
                    // out as a black rectangle here: blind-J ja/educational__2aa00bc4's
                    // turned colour cards each carried an upright black frame.
                    let lw = (*stroke_width as f64).max(0.5);
                    let sc = stroke_color.as_deref()
                        .and_then(|c| parse_hex_color(c))
                        .unwrap_or(Color::Rgb(0.0, 0.0, 0.0));
                    let stroke = Some(oxipdf_core::ir::StrokeStyle {
                        color: sc.clone(),
                        width: lw,
                        line_cap: oxipdf_core::ir::LineCap::Butt,
                        line_join: oxipdf_core::ir::LineJoin::Miter,
                    });
                    let line = |ops: Vec<PathOp>| ContentElement::Path(PathData {
                        operations: ops, stroke: stroke.clone(), fill: None,
                    });
                    // the renderer's arrowhead: length max(3w, 3.75pt), half width max(1.8w, 2.25pt)
                    let arrow = |tx: f64, ty: f64, dx: f64, dy: f64| -> Option<ContentElement> {
                        let len = (dx * dx + dy * dy).sqrt();
                        if len < 0.01 { return None; }
                        let (ux, uy) = (dx / len, dy / len);
                        let (px, py) = (-uy, ux);
                        let alen = (lw * 3.0).max(3.75);
                        let ahw = (lw * 1.8).max(2.25);
                        let (bx, by) = (tx - ux * alen, ty - uy * alen);
                        Some(ContentElement::Path(PathData {
                            operations: vec![
                                PathOp::MoveTo(tx, ty),
                                PathOp::LineTo(bx + px * ahw, by + py * ahw),
                                PathOp::LineTo(bx - px * ahw, by - py * ahw),
                                PathOp::ClosePath,
                            ],
                            stroke: None,
                            fill: Some(sc.clone()),
                        }))
                    };
                    match shape_type.as_str() {
                        "line" | "straightConnector1" | "bentConnector3" | "bentConnector2" => {
                            let (sx, sy) = (if *flip_h { x + ew } else { x }, if *flip_v { y + eh } else { y });
                            let (ex, ey) = (if *flip_h { x } else { x + ew }, if *flip_v { y } else { y + eh });
                            let (tf, hf);
                            if shape_type.starts_with("bentConnector") {
                                let mx = (sx + ex) * 0.5;
                                contents.push(line(vec![PathOp::MoveTo(sx, sy), PathOp::LineTo(mx, sy),
                                    PathOp::LineTo(mx, ey), PathOp::LineTo(ex, ey)]));
                                tf = (mx, ey); hf = (mx, sy);
                            } else {
                                contents.push(line(vec![PathOp::MoveTo(sx, sy), PathOp::LineTo(ex, ey)]));
                                tf = (sx, sy); hf = (ex, ey);
                            }
                            if *arrow_tail {
                                contents.extend(arrow(ex, ey, ex - tf.0, ey - tf.1));
                            }
                            if *arrow_head {
                                contents.extend(arrow(sx, sy, sx - hf.0, sy - hf.1));
                            }
                        }
                        "ellipse" | "oval" => {
                            let (cx, cy, rx, ry) = (x + ew / 2.0, y + eh / 2.0, ew / 2.0, eh / 2.0);
                            let (kx, ky) = (rx * 0.5523, ry * 0.5523);
                            contents.push(line(vec![
                                PathOp::MoveTo(cx + rx, cy),
                                PathOp::CurveTo(cx + rx, cy + ky, cx + kx, cy + ry, cx, cy + ry),
                                PathOp::CurveTo(cx - kx, cy + ry, cx - rx, cy + ky, cx - rx, cy),
                                PathOp::CurveTo(cx - rx, cy - ky, cx - kx, cy - ry, cx, cy - ry),
                                PathOp::CurveTo(cx + kx, cy - ry, cx + rx, cy - ky, cx + rx, cy),
                                PathOp::ClosePath,
                            ]));
                        }
                        "bracketPair" | "leftBracket" | "rightBracket" => {
                            let r = eh.min(ew) * 16667.0 / 100000.0;
                            let k = r * 0.5523;
                            if shape_type != "rightBracket" {
                                contents.push(line(vec![
                                    PathOp::MoveTo(x + r, y),
                                    PathOp::CurveTo(x + r - k, y, x, y + r - k, x, y + r),
                                    PathOp::LineTo(x, y + eh - r),
                                    PathOp::CurveTo(x, y + eh - r + k, x + r - k, y + eh, x + r, y + eh),
                                ]));
                            }
                            if shape_type != "leftBracket" {
                                contents.push(line(vec![
                                    PathOp::MoveTo(x + ew - r, y),
                                    PathOp::CurveTo(x + ew - r + k, y, x + ew, y + r - k, x + ew, y + r),
                                    PathOp::LineTo(x + ew, y + eh - r),
                                    PathOp::CurveTo(x + ew, y + eh - r + k, x + ew - r + k, y + eh, x + ew - r, y + eh),
                                ]));
                            }
                        }
                        _ => {}
                    }
                }
                oxidocs_core::layout::LayoutContent::PresetShape { ref shape_type, ref stroke_color, stroke_width, .. } => {
                    let sx = x;
                    let sy = y;
                    let sw = ew;
                    let sh = eh;
                    let lw = *stroke_width as f64;
                    let sc = stroke_color.as_deref()
                        .and_then(|c| parse_hex_color(c))
                        .unwrap_or(Color::Rgb(0.0, 0.0, 0.0));
                    let stroke_style = Some(oxipdf_core::ir::StrokeStyle {
                        color: sc,
                        width: lw,
                        line_cap: oxipdf_core::ir::LineCap::Butt,
                        line_join: oxipdf_core::ir::LineJoin::Miter,
                    });

                    match shape_type.as_str() {
                        "bracketPair" => {
                            // OOXML bracketPair: left 〔 and right 〕 brackets
                            // Default adjustment value = 16667/100000 ≈ 1/6 of min dimension
                            let r = sh.min(sw) * 16667.0 / 100000.0;
                            // Bezier constant for quarter-circle approximation
                            let k = r * 0.5523;
                            // Left bracket 〔
                            contents.push(ContentElement::Path(PathData {
                                operations: vec![
                                    PathOp::MoveTo(sx + r, sy),
                                    PathOp::CurveTo(sx + r - k, sy, sx, sy + r - k, sx, sy + r),
                                    PathOp::LineTo(sx, sy + sh - r),
                                    PathOp::CurveTo(sx, sy + sh - r + k, sx + r - k, sy + sh, sx + r, sy + sh),
                                ],
                                stroke: stroke_style.clone(),
                                fill: None,
                            }));
                            // Right bracket 〕
                            contents.push(ContentElement::Path(PathData {
                                operations: vec![
                                    PathOp::MoveTo(sx + sw - r, sy),
                                    PathOp::CurveTo(sx + sw - r + k, sy, sx + sw, sy + r - k, sx + sw, sy + r),
                                    PathOp::LineTo(sx + sw, sy + sh - r),
                                    PathOp::CurveTo(sx + sw, sy + sh - r + k, sx + sw - r + k, sy + sh, sx + sw - r, sy + sh),
                                ],
                                stroke: stroke_style,
                                fill: None,
                            }));
                        }
                        _ => {
                            // Generic: draw outline rectangle
                            contents.push(ContentElement::Path(PathData {
                                operations: vec![
                                    PathOp::MoveTo(sx, sy),
                                    PathOp::LineTo(sx + sw, sy),
                                    PathOp::LineTo(sx + sw, sy + sh),
                                    PathOp::LineTo(sx, sy + sh),
                                    PathOp::ClosePath,
                                ],
                                stroke: stroke_style,
                                fill: None,
                            }));
                        }
                    }
                }
                // R-05a: balloon variants are layout-only stubs until R-05g
                // wires the PDF rendering. Skipping is correct for now;
                // PDF output omits comment balloons entirely (matches the
                // pre-R-05 state).
                oxidocs_core::layout::LayoutContent::Balloon { .. } => {}
                oxidocs_core::layout::LayoutContent::BalloonConnector { .. } => {}
                // S1120 custGeom paths: not surfaced in the CLI page model yet
                oxidocs_core::layout::LayoutContent::VectorPath { segs, fill, stroke_color, stroke_width } => {
                    // A group member's outline (custGeom curves, presets, SmartArt
                    // shapes), its points normalized 0..1 over the element box.
                    // This arm drew nothing, so every such outline was missing
                    // from docx-to-pdf while the screen renderers drew it.
                    use oxidocs_core::ir::PathSeg;
                    let px = |u: f32| x + u as f64 * ew;
                    let py = |v: f32| y + v as f64 * eh;
                    let ops: Vec<PathOp> = segs.iter().map(|seg| match *seg {
                        PathSeg::M(a, b) => PathOp::MoveTo(px(a), py(b)),
                        PathSeg::L(a, b) => PathOp::LineTo(px(a), py(b)),
                        PathSeg::C(a, b, c, d, e, f) => PathOp::CurveTo(px(a), py(b), px(c), py(d), px(e), py(f)),
                        PathSeg::Z => PathOp::ClosePath,
                    }).collect();
                    let fill_color = fill.as_deref().and_then(parse_hex_color);
                    let stroke = stroke_color.as_deref().and_then(parse_hex_color)
                        .filter(|_| *stroke_width > 0.0)
                        .map(|color| StrokeStyle {
                            color, width: *stroke_width as f64, line_cap: LineCap::Butt, line_join: LineJoin::Round,
                        });
                    if !ops.is_empty() && (fill_color.is_some() || stroke.is_some()) {
                        contents.push(ContentElement::Path(PathData { operations: ops, stroke, fill: fill_color }));
                    }
                }
            }
        }

        Page {
            width: w,
            height: h,
            media_box: Rectangle { llx: 0.0, lly: 0.0, urx: w, ury: h },
            crop_box: None,
            contents,
            rotation: 0,
        }
    }).collect();

    let title = doc.pages.first()
        .and_then(|p| p.blocks.first())
        .and_then(|b| match b {
            Block::Paragraph(p) => {
                let text: String = p.runs.iter().map(|r| r.text.as_str()).collect();
                if !text.is_empty() { Some(text) } else { None }
            }
            _ => None,
        });

    let pdf = PdfDocument {
        version: PdfVersion::new(1, 7),
        info: DocumentInfo {
            title,
            producer: Some("Oxi (oxipdf-core)".to_string()),
            ..Default::default()
        },
        pages,
        outline: Vec::new(),
        embedded_fonts,
    };
    (pdf, layout_result.hf_parts_shown)
}

/// Render cover page elements with professional layout
fn render_cover_element(builder: &mut PdfBuilder, para: &Paragraph, role: ParaRole, x_offset: f64, available_width: f64) {
    let text: String = para.runs.iter().map(|r| r.text.as_str()).collect();
    if text.trim().is_empty() { return; }

    let font_family = para.runs.first()
        .and_then(|r| r.style.font_family.as_deref().or(r.style.font_family_east_asia.as_deref()));
    let color_hex = para.runs.first().and_then(|r| r.style.color.as_deref());
    let color = color_hex.and_then(parse_hex_color).unwrap_or(Color::Rgb(0.18, 0.25, 0.34));

    match role {
        ParaRole::CoverTitle => {
            // Auto-size: try 20pt, reduce until it fits on 1 line (or 2 max)
            let mut font_size = 20.0;
            loop {
                let lines = builder.wrap_text_accurate(&text, font_size, available_width);
                if lines.len() <= 1 || font_size <= 14.0 {
                    break;
                }
                font_size -= 0.5;
            }
            let font_name = resolve_font(font_family, true);
            let line_height = font_size * 1.6;

            // Position title at approximately 1/3 from top of page (top-down coords)
            let target_y = builder.page_height * 0.28;
            if builder.cursor_y < target_y {
                builder.cursor_y = target_y;
            }

            let lines = builder.wrap_text_accurate(&text, font_size, available_width);
            for line in &lines {
                let line_w = builder.text_width(line, font_size);
                let x = x_offset + (available_width - line_w) / 2.0;
                builder.add_text(x, builder.cursor_y, line.clone(), font_name.clone(), font_size, color);
                builder.cursor_y += line_height;
            }

            // Decorative line under title
            builder.cursor_y += 6.0;
            builder.add_line(
                x_offset + available_width * 0.2, builder.cursor_y,
                x_offset + available_width * 0.8, builder.cursor_y,
                1.5, color,
            );
            builder.cursor_y += 20.0;
        }
        ParaRole::CoverSubtitle => {
            let font_size = 12.0;
            let font_name = resolve_font(font_family, true);
            let line_w = builder.text_width(&text, font_size);
            let x = x_offset + (available_width - line_w) / 2.0;
            builder.add_text(x, builder.cursor_y, text, font_name, font_size, Color::Gray(0.3));
            builder.cursor_y += font_size * 3.0;
        }
        ParaRole::CoverAuthor => {
            let font_size = 11.0;
            let font_name = resolve_font(font_family, false);
            let line_w = builder.text_width(&text, font_size);
            let x = x_offset + (available_width - line_w) / 2.0;
            builder.add_text(x, builder.cursor_y, text, font_name, font_size, Color::Gray(0.25));
            builder.cursor_y += font_size * 2.0;
        }
        _ => {}
    }
}

/// Render a paragraph with role-based styling
fn render_paragraph_styled(builder: &mut PdfBuilder, para: &Paragraph, role: ParaRole, x_offset: f64, available_width: f64, registry: &FontMetricsRegistry, doc_default_font: Option<&str>, doc_default_font_size: f32, grid_pitch_pt: f64) {
    let full_text: String = para.runs.iter().map(|r| r.text.as_str()).collect();

    // Empty paragraph — add spacing based on default font metrics
    if full_text.trim().is_empty() && para.style.list_marker.is_none() {
        let default_fs = para.style.default_run_style.as_ref()
            .and_then(|rs| rs.font_size).unwrap_or(10.5);
        let default_ff = para.style.default_run_style.as_ref()
            .and_then(|rs| rs.font_family.as_deref());
        let lh = compute_line_height_with_default(
            registry, default_ff, default_fs,
            para.style.line_spacing, para.style.line_spacing_rule.as_deref(),
            doc_default_font, Some(doc_default_font_size), grid_pitch_pt, grid_pitch_pt > 0.01,
        );
        let gap = para.style.space_before.unwrap_or(0.0) as f64
            + lh
            + para.style.space_after.unwrap_or(0.0) as f64;
        builder.cursor_y += gap;
        return;
    }

    // Determine font size from run styling
    let default_size = para.runs.first()
        .and_then(|r| r.style.font_size)
        .or(para.style.default_run_style.as_ref().and_then(|rs| rs.font_size))
        .unwrap_or(10.5) as f64;

    let default_font = para.runs.first()
        .and_then(|r| r.style.font_family.as_deref().or(r.style.font_family_east_asia.as_deref()))
        .or(para.style.default_run_style.as_ref().and_then(|rs| rs.font_family.as_deref()));

    let _default_color_hex = para.runs.first()
        .and_then(|r| r.style.color.as_deref())
        .or(para.style.default_run_style.as_ref().and_then(|rs| rs.color.as_deref()));

    // Role-based styling overrides
    match role {
        ParaRole::SectionHeader => {
            render_section_header(builder, para, x_offset, available_width, registry, doc_default_font, grid_pitch_pt);
            return;
        }
        ParaRole::SubHeader => {
            render_sub_header(builder, para, x_offset, available_width, registry, doc_default_font, grid_pitch_pt);
            return;
        }
        ParaRole::SubSubHeader => {
            render_sub_sub_header(builder, para, x_offset, available_width, registry, doc_default_font, grid_pitch_pt);
            return;
        }
        ParaRole::CreditFooter => {
            render_credit_footer(builder, para, x_offset, available_width, registry, doc_default_font, grid_pitch_pt);
            return;
        }
        _ => {} // Body and ListItem handled below
    }

    // --- Regular body / list rendering ---
    let font_size = default_size;
    let line_height = compute_line_height_with_default(
        registry,
        default_font,
        font_size as f32,
        para.style.line_spacing,
        para.style.line_spacing_rule.as_deref(),
        doc_default_font,
        Some(doc_default_font_size),
        grid_pitch_pt,
        grid_pitch_pt > 0.01,
    );

    let space_before = para.style.space_before.unwrap_or(0.0) as f64;
    let space_after = para.style.space_after.unwrap_or(0.0) as f64;

    // Indentation
    let indent_left = para.style.indent_left.unwrap_or(0.0) as f64;
    let indent_first = para.style.indent_first_line.unwrap_or(0.0) as f64;

    builder.cursor_y += space_before;

    if builder.needs_page_break(line_height + 10.0) {
        builder.new_page();
    }

    // Paragraph shading (background fill)
    if let Some(ref shd) = para.style.shading {
        let lines_est = builder.wrap_text_accurate(&full_text, font_size, available_width - indent_left).len();
        let para_h = lines_est as f64 * line_height;
        render_paragraph_shading(builder, shd, x_offset, builder.cursor_y, available_width, para_h);
    }

    // Paragraph borders
    if let Some(ref borders) = para.style.borders {
        let lines_est = builder.wrap_text_accurate(&full_text, font_size, available_width - indent_left).len();
        let para_h = lines_est as f64 * line_height;
        render_paragraph_borders(builder, borders, x_offset, builder.cursor_y, available_width, para_h);
    }

    // List marker handling (both IR list_marker and text-based ・)
    let mut marker_offset = 0.0;
    let mut effective_indent_left = indent_left;

    if para.style.list_marker.is_some() || role == ParaRole::ListItem {
        if let Some(marker) = &para.style.list_marker {
            let marker_indent = para.style.list_indent.unwrap_or(18.0) as f64;
            marker_offset = marker_indent;
            let marker_x = x_offset + indent_left;
            let marker_font = resolve_font(default_font, false);
            // Marker baseline matches first text line (cursor_y + line_height)
            builder.add_text(
                marker_x, builder.cursor_y + line_height,
                marker.clone(), marker_font, font_size, Color::Gray(0.0),
            );
        } else if full_text.trim().starts_with('・') || full_text.trim().starts_with('※') {
            effective_indent_left = effective_indent_left.max(8.0);
        }
    }

    let text_x_start = x_offset + effective_indent_left + marker_offset;
    let text_width = (available_width - effective_indent_left - marker_offset).max(50.0);

    let para_default_bold = para.style.default_run_style.as_ref().map(|rs| rs.bold).unwrap_or(false);
    let para_default_ff = para.style.default_run_style.as_ref().and_then(|rs| rs.font_family.as_deref());
    let lines = wrap_runs_into_lines(&para.runs, font_size, text_width, indent_first, &builder.font_widths, para_default_bold, para_default_ff);

    for (line_idx, line_runs) in lines.iter().enumerate() {
        if builder.needs_page_break(line_height) {
            builder.new_page();
        }

        let line_x = if line_idx == 0 {
            text_x_start + indent_first
        } else {
            text_x_start
        };

        let line_width: f64 = line_runs.iter()
            .map(|lr| {
                let base_w = builder.text_width(&lr.text, lr.font_size);
                let spacing_w = lr.char_spacing * lr.text.chars().count() as f64;
                base_w + spacing_w
            })
            .sum();

        let align_offset = match para.alignment {
            Alignment::Center => (text_width - line_width) / 2.0,
            Alignment::Right => text_width - line_width,
            _ => 0.0,
        };

        // Advance cursor FIRST (Word places baseline at cursor + line_height)
        builder.cursor_y += line_height;
        let text_y = builder.cursor_y;

        let mut run_x = line_x + align_offset.max(0.0);

        // Distribute alignment: spread characters evenly across the line
        if para.alignment == Alignment::Distribute && line_width > 0.01 {
            let total_chars: usize = line_runs.iter().map(|lr| lr.text.chars().count()).sum();
            if total_chars > 1 {
                let slack = text_width - line_width;
                let extra_per_gap = slack / (total_chars - 1) as f64;
                for lr in line_runs {
                    for ch in lr.text.chars() {
                        let ch_str = ch.to_string();
                        builder.add_text(
                            run_x, text_y,
                            ch_str.clone(), lr.font_name.clone(), lr.font_size, lr.color,
                        );
                        run_x += builder.text_width(&ch_str, lr.font_size) + extra_per_gap;
                    }
                }
            } else {
                for lr in line_runs {
                    if !lr.text.is_empty() {
                        builder.add_text(
                            run_x, text_y,
                            lr.text.clone(), lr.font_name.clone(), lr.font_size, lr.color,
                        );
                        run_x += builder.text_width(&lr.text, lr.font_size);
                    }
                }
            }
        } else {
            for lr in line_runs {
                if !lr.text.is_empty() {
                    let run_start_x = run_x;
                    if lr.char_spacing.abs() > 0.01 {
                        // Render character by character with extra spacing
                        for ch in lr.text.chars() {
                            let ch_str = ch.to_string();
                            builder.add_text(
                                run_x, text_y,
                                ch_str.clone(), lr.font_name.clone(), lr.font_size, lr.color,
                            );
                            run_x += builder.text_width(&ch_str, lr.font_size) + lr.char_spacing;
                        }
                    } else {
                        builder.add_text(
                            run_x, text_y,
                            lr.text.clone(), lr.font_name.clone(), lr.font_size, lr.color,
                        );
                        run_x += builder.text_width(&lr.text, lr.font_size);
                    }
                    // Draw underline
                    if lr.underline {
                        let ul_y = text_y + lr.font_size * 0.15;
                        let ul_w = 0.5_f64.max(lr.font_size * 0.05);
                        let is_double = lr.underline_style.as_deref() == Some("double");
                        if is_double {
                            // Double underline: two lines with gap
                            let gap = ul_w * 2.0;
                            builder.add_line(run_start_x, ul_y - gap / 2.0, run_x, ul_y - gap / 2.0, ul_w, lr.color);
                            builder.add_line(run_start_x, ul_y + gap / 2.0, run_x, ul_y + gap / 2.0, ul_w, lr.color);
                        } else {
                            builder.add_line(run_start_x, ul_y, run_x, ul_y, ul_w, lr.color);
                        }
                    }
                }
            }
        }
    }

    builder.cursor_y += space_after;
}

/// Extract paragraph font properties
fn para_font_props(para: &Paragraph) -> (f64, Option<&str>, bool) {
    let font_size = para.runs.first()
        .and_then(|r| r.style.font_size)
        .or(para.style.default_run_style.as_ref().and_then(|rs| rs.font_size))
        .unwrap_or(10.5) as f64;
    let font_family = para.runs.first()
        .and_then(|r| r.style.font_family.as_deref().or(r.style.font_family_east_asia.as_deref()))
        .or(para.style.default_run_style.as_ref().and_then(|rs| rs.font_family.as_deref()));
    let bold = para.runs.first().map(|r| r.style.bold).unwrap_or(false)
        || para.style.default_run_style.as_ref().map(|rs| rs.bold).unwrap_or(false);
    (font_size, font_family, bold)
}

fn para_color(para: &Paragraph) -> Color {
    para.runs.first()
        .and_then(|r| r.style.color.as_deref())
        .or(para.style.default_run_style.as_ref().and_then(|rs| rs.color.as_deref()))
        .and_then(parse_hex_color)
        .unwrap_or(Color::Gray(0.0))
}

/// Render ① ② ③ section headers — uses actual paragraph properties
fn render_section_header(builder: &mut PdfBuilder, para: &Paragraph, x_offset: f64, available_width: f64, registry: &FontMetricsRegistry, doc_default_font: Option<&str>, grid_pitch_pt: f64) {
    let text: String = para.runs.iter().map(|r| r.text.as_str()).collect();
    let (font_size, font_family, bold) = para_font_props(para);
    let line_height = compute_line_height(
        registry, font_family, font_size as f32,
        para.style.line_spacing, para.style.line_spacing_rule.as_deref(),
        doc_default_font, grid_pitch_pt,
    );
    let space_before = para.style.space_before.unwrap_or(0.0) as f64;
    let space_after = para.style.space_after.unwrap_or(0.0) as f64;
    let font_name = resolve_font(font_family, bold);
    let color = para_color(para);

    builder.cursor_y += space_before;
    if builder.needs_page_break(line_height + 10.0) { builder.new_page(); }

    // Shading background if present
    if let Some(ref shd) = para.style.shading {
        if let Some(bg) = parse_hex_color(shd) {
            builder.add_rect_fill(x_offset, builder.cursor_y - 1.0, available_width, line_height + 2.0, bg);
        }
    }

    let indent_left = para.style.indent_left.unwrap_or(0.0) as f64;
    add_text_split(builder, x_offset + indent_left, builder.cursor_y, text, font_name, font_size, color);
    builder.cursor_y += line_height + space_after;
}

/// Render "N-N." sub-headers — uses actual paragraph properties
fn render_sub_header(builder: &mut PdfBuilder, para: &Paragraph, x_offset: f64, _available_width: f64, registry: &FontMetricsRegistry, doc_default_font: Option<&str>, grid_pitch_pt: f64) {
    let text: String = para.runs.iter().map(|r| r.text.as_str()).collect();
    let (font_size, font_family, bold) = para_font_props(para);
    let line_height = compute_line_height(
        registry, font_family, font_size as f32,
        para.style.line_spacing, para.style.line_spacing_rule.as_deref(),
        doc_default_font, grid_pitch_pt,
    );
    let space_before = para.style.space_before.unwrap_or(0.0) as f64;
    let space_after = para.style.space_after.unwrap_or(0.0) as f64;
    let font_name = resolve_font(font_family, bold);
    let color = para_color(para);

    builder.cursor_y += space_before;
    if builder.needs_page_break(line_height + 10.0) { builder.new_page(); }

    let indent_left = para.style.indent_left.unwrap_or(0.0) as f64;
    add_text_split(builder, x_offset + indent_left, builder.cursor_y, text, font_name, font_size, color);
    builder.cursor_y += line_height + space_after;
}

/// Render sub-sub-headers — uses actual paragraph properties
fn render_sub_sub_header(builder: &mut PdfBuilder, para: &Paragraph, x_offset: f64, available_width: f64, registry: &FontMetricsRegistry, doc_default_font: Option<&str>, grid_pitch_pt: f64) {
    let text: String = para.runs.iter().map(|r| r.text.as_str()).collect();
    let (font_size, font_family, bold) = para_font_props(para);
    let line_height = compute_line_height(
        registry, font_family, font_size as f32,
        para.style.line_spacing, para.style.line_spacing_rule.as_deref(),
        doc_default_font, grid_pitch_pt,
    );
    let space_before = para.style.space_before.unwrap_or(0.0) as f64;
    let space_after = para.style.space_after.unwrap_or(0.0) as f64;
    let font_name = resolve_font(font_family, bold);
    let color = para_color(para);

    builder.cursor_y += space_before;
    if builder.needs_page_break(line_height + 10.0) { builder.new_page(); }

    let indent_left = para.style.indent_left.unwrap_or(0.0) as f64;
    let lines = wrap_text(&text, font_size, available_width - indent_left);
    for line in &lines {
        add_text_split(builder, x_offset + indent_left, builder.cursor_y, line.clone(), font_name.clone(), font_size, color);
        builder.cursor_y += line_height;
    }
    builder.cursor_y += space_after;
}

/// Render the credit footer line — uses actual paragraph properties
fn render_credit_footer(builder: &mut PdfBuilder, para: &Paragraph, x_offset: f64, available_width: f64, registry: &FontMetricsRegistry, doc_default_font: Option<&str>, grid_pitch_pt: f64) {
    let text: String = para.runs.iter().map(|r| r.text.as_str()).collect();
    let (font_size, font_family, bold) = para_font_props(para);
    let line_height = compute_line_height(
        registry, font_family, font_size as f32,
        para.style.line_spacing, para.style.line_spacing_rule.as_deref(),
        doc_default_font, grid_pitch_pt,
    );
    let space_before = para.style.space_before.unwrap_or(0.0) as f64;
    let space_after = para.style.space_after.unwrap_or(0.0) as f64;
    let font_name = resolve_font(font_family, bold);
    let color = para_color(para);

    builder.cursor_y += space_before;

    let text_w = builder.text_width(&text, font_size);
    let x = match para.alignment {
        Alignment::Right => x_offset + available_width - text_w,
        Alignment::Center => x_offset + (available_width - text_w) / 2.0,
        _ => x_offset,
    };
    add_text_split(builder, x, builder.cursor_y, text, font_name, font_size, color);
    builder.cursor_y += line_height + space_after;
}

/// Resolve the CJK font name for non-ASCII text.
/// If font_name is a known CJK font (MS Gothic, MS Mincho, etc.), use it directly.
/// Otherwise, fall back to "OxiCJK-Regular"/"OxiCJK-Bold".
fn resolve_cjk_font(font_name: &str, bold: bool) -> String {
    // Check if the font_name is already a CJK font that we embed per-family
    if font_name.contains("Gothic") || font_name.contains("Mincho") || font_name.contains("ゴシック") || font_name.contains("明朝")
        || font_name.contains("游") || font_name.contains("ＭＳ")
        || font_name.contains("メイリオ") || font_name.contains("Meiryo")
    {
        return font_name.to_string();
    }
    if bold { "OxiCJK-Bold".to_string() } else { "OxiCJK-Regular".to_string() }
}

/// Emit text with proper ASCII/CJK font splitting.
/// ASCII characters use the given `font_name` (Type1), non-ASCII use the resolved CJK font.
/// When font_name IS a CJK font, all characters use that font (no splitting).
fn add_text_split(builder: &mut PdfBuilder, x: f64, y: f64, text: String, font_name: String, font_size: f64, color: Color) {
    // If font_name is a CJK font, use it for ALL characters (including ASCII)
    let is_cjk_font = font_name.contains("Gothic") || font_name.contains("Mincho") || font_name.contains("ゴシック") || font_name.contains("明朝")
        || font_name.contains("游") || font_name.contains("ＭＳ")
        || font_name.contains("メイリオ") || font_name.contains("Meiryo");
    if is_cjk_font {
        builder.add_text(x, y, text, font_name, font_size, color);
        return;
    }

    if !text.chars().any(|c| c as u32 > 0x7F) {
        // Pure ASCII — use the regular font directly
        builder.add_text(x, y, text, font_name, font_size, color);
        return;
    }

    let cjk_font = resolve_cjk_font(&font_name, font_name.contains("Bold"));

    // Split text into ASCII and non-ASCII segments
    let mut cur_x = x;
    let mut buf = String::new();
    let mut buf_is_cjk = false;

    for ch in text.chars() {
        let ch_is_cjk = ch as u32 > 0x7F;
        if !buf.is_empty() && ch_is_cjk != buf_is_cjk {
            let fn_for_buf = if buf_is_cjk { cjk_font.clone() } else { font_name.clone() };
            let w = builder.text_width(&buf, font_size);
            builder.add_text(cur_x, y, buf.clone(), fn_for_buf, font_size, color);
            cur_x += w;
            buf.clear();
        }
        buf_is_cjk = ch_is_cjk;
        buf.push(ch);
    }
    if !buf.is_empty() {
        let fn_for_buf = if buf_is_cjk { cjk_font } else { font_name };
        builder.add_text(cur_x, y, buf, fn_for_buf, font_size, color);
    }
}

/// A styled text fragment for a single line
struct LineRun {
    text: String,
    font_name: String,
    font_size: f64,
    color: Color,
    char_spacing: f64,
    underline: bool,
    underline_style: Option<String>,
}

/// Break runs into wrapped lines, preserving per-run styling.
/// Uses embedded font widths (font_widths) for accurate character width measurement
/// that matches the PDF renderer's actual glyph widths.
fn wrap_runs_into_lines(
    runs: &[oxidocs_core::ir::Run],
    default_font_size: f64,
    max_width: f64,
    first_line_indent: f64,
    font_widths: &HashMap<u32, u16>,
    default_bold: bool,
    default_font_family: Option<&str>,
) -> Vec<Vec<LineRun>> {
    let mut lines: Vec<Vec<LineRun>> = Vec::new();
    let mut current_line: Vec<LineRun> = Vec::new();
    let mut current_width = first_line_indent;
    let effective_max = max_width;

    for run in runs {
        let font_size = run.style.font_size.unwrap_or(default_font_size as f32) as f64;
        let bold = run.style.bold || default_bold;
        let font_family = run.style.font_family.as_deref()
            .or(run.style.font_family_east_asia.as_deref())
            .or(default_font_family);
        let font_name = resolve_font(font_family, bold);
        let color = run.style.color.as_deref()
            .and_then(parse_hex_color)
            .unwrap_or(Color::Gray(0.0));
        // Word doubles the character spacing (applies on both sides of each character)
        let char_spacing = run.style.character_spacing.unwrap_or(0.0) as f64 * 2.0;
        let underline = run.style.underline;
        let underline_style = run.style.underline_style.clone();

        // Process text character by character for line wrapping.
        // Split into ASCII (Type1 base font) and non-ASCII (CJK CIDFont) segments.
        let mut buf = String::new();
        let mut buf_width = 0.0;
        let mut buf_is_cjk = false;

        for ch in run.text.chars() {
            if ch == '\t' {
                // Tab: flush buffer, then advance to next tab stop
                if !buf.is_empty() {
                    let fn_for_buf = if buf_is_cjk { "OxiCJK-Regular".to_string() } else { font_name.clone() };
                    current_line.push(LineRun {
                        text: buf.clone(), font_name: fn_for_buf,
                        font_size, color, char_spacing, underline, underline_style: underline_style.clone(),
                    });
                    current_width += buf_width;
                    buf.clear();
                    buf_width = 0.0;
                }
                // Advance to next default tab stop (Word default: every 36pt = 0.5 inch)
                let tab_interval = 36.0;
                let next_tab = ((current_width / tab_interval).floor() + 1.0) * tab_interval;
                let tab_gap = (next_tab - current_width).max(font_size * 0.25);
                // Approximate tab with spaces
                let space_w = font_size * 0.25;
                let num_spaces = (tab_gap / space_w).round().max(1.0) as usize;
                let tab_text = " ".repeat(num_spaces);
                current_line.push(LineRun {
                    text: tab_text, font_name: font_name.clone(),
                    font_size, color, char_spacing: 0.0, underline: false, underline_style: None,
                });
                current_width = next_tab;
                continue;
            }
            if ch == '\n' || ch == '\r' {
                // Explicit line break
                if !buf.is_empty() {
                    let fn_for_buf = if buf_is_cjk { "OxiCJK-Regular".to_string() } else { font_name.clone() };
                    current_line.push(LineRun {
                        text: buf.clone(), font_name: fn_for_buf,
                        font_size, color, char_spacing, underline, underline_style: underline_style.clone(),
                    });
                    buf.clear();
                    buf_width = 0.0;
                }
                lines.push(std::mem::take(&mut current_line));
                current_width = 0.0;
                continue;
            }

            let ch_is_cjk = ch as u32 > 0x7F;

            // If script changes (ASCII↔CJK), flush buffer as a separate run
            if !buf.is_empty() && ch_is_cjk != buf_is_cjk {
                let fn_for_buf = if buf_is_cjk { "OxiCJK-Regular".to_string() } else { font_name.clone() };
                current_line.push(LineRun {
                    text: buf.clone(), font_name: fn_for_buf,
                    font_size, color, char_spacing, underline, underline_style: underline_style.clone(),
                });
                current_width += buf_width;
                buf.clear();
                buf_width = 0.0;
            }
            buf_is_cjk = ch_is_cjk;

            // Use embedded font widths when available, matching PDF renderer
            let ch_width = (if let Some(&glyph_w) = font_widths.get(&(ch as u32)) {
                font_size * (glyph_w as f64 / 1000.0)
            } else {
                char_width(ch, font_size)
            }) + char_spacing;

            if current_width + buf_width + ch_width > effective_max && !(current_line.is_empty() && buf.is_empty()) {
                // Wrap: flush buffer to current line, start new line
                if !buf.is_empty() {
                    let fn_for_buf = if buf_is_cjk { "OxiCJK-Regular".to_string() } else { font_name.clone() };
                    current_line.push(LineRun {
                        text: buf.clone(), font_name: fn_for_buf,
                        font_size, color, char_spacing, underline, underline_style: underline_style.clone(),
                    });
                    buf.clear();
                    buf_width = 0.0;
                }
                lines.push(std::mem::take(&mut current_line));
                current_width = 0.0;
            }

            buf.push(ch);
            buf_width += ch_width;
        }

        // Flush remaining buffer
        if !buf.is_empty() {
            let fn_for_buf = if buf_is_cjk { "OxiCJK-Regular".to_string() } else { font_name.clone() };
            current_line.push(LineRun {
                text: buf, font_name: fn_for_buf,
                font_size, color, char_spacing, underline, underline_style: underline_style.clone(),
            });
            current_width += buf_width;
        }
    }

    // Flush last line
    if !current_line.is_empty() {
        lines.push(current_line);
    }

    if lines.is_empty() {
        lines.push(Vec::new());
    }

    lines
}

fn render_table(builder: &mut PdfBuilder, table: &Table, x_offset: f64, available_width: f64, registry: &FontMetricsRegistry, doc_default_font: Option<&str>) {
    // Build per-grid-column widths from tblGrid (canonical), cell widths, or autofit.
    // Use actual widths from the docx — do NOT scale to available_width.
    let grid_col_widths: Vec<f64> = if !table.grid_columns.is_empty() {
        table.grid_columns.iter().map(|&w| w as f64).collect()
    } else if let Some(first_row) = table.rows.first() {
        let specified: Vec<Option<f64>> = first_row.cells.iter()
            .map(|c| c.width.map(|w| w as f64))
            .collect();
        let total_specified: f64 = specified.iter().filter_map(|w| *w).sum();
        let num_cols = first_row.cells.len();
        if total_specified > 0.0 {
            specified.iter().map(|w| w.unwrap_or(total_specified / num_cols as f64)).collect()
        } else {
            // AutoFit to Contents: column width = max text width + cell padding
            let pad_lr = 5.4 * 2.0; // default L+R cell padding (10.8pt)
            let min_col_w = 11.1; // empty cell minimum (padding only)
            let mut col_widths = vec![0.0_f64; num_cols];
            for row in &table.rows {
                for (ci, cell) in row.cells.iter().enumerate() {
                    if ci >= num_cols { break; }
                    let mut max_w = 0.0_f64;
                    for block in &cell.blocks {
                        if let Block::Paragraph(para) = block {
                            // Measure text width using font metrics registry
                            for run in &para.runs {
                                let fs = run.style.font_size.unwrap_or(10.5);
                                let font_family = run.style.font_family.as_deref()
                                    .or(doc_default_font)
                                    .unwrap_or("Calibri");
                                let metrics = registry.get(font_family);
                                let w: f64 = run.text.chars()
                                    .map(|c| metrics.char_width_pt(c, fs) as f64)
                                    .sum();
                                max_w += w;
                            }
                        }
                    }
                    col_widths[ci] = col_widths[ci].max(max_w + pad_lr);
                }
            }
            // Ensure minimum column width for empty cells
            for w in &mut col_widths {
                if *w < min_col_w { *w = min_col_w; }
            }
            col_widths
        }
    } else {
        return;
    };

    // Default cell margins from table style (Word default: 5.4pt left/right, 0 top/bottom)
    let cell_margin_l = table.style.default_cell_margins.as_ref()
        .and_then(|m| m.left).unwrap_or(5.4) as f64;
    let cell_margin_r = table.style.default_cell_margins.as_ref()
        .and_then(|m| m.right).unwrap_or(5.4) as f64;
    let cell_margin_t = table.style.default_cell_margins.as_ref()
        .and_then(|m| m.top).unwrap_or(0.0) as f64;
    let cell_margin_b = table.style.default_cell_margins.as_ref()
        .and_then(|m| m.bottom).unwrap_or(0.0) as f64;

    // Table alignment: shift x_offset based on alignment + indent
    let table_total_width: f64 = grid_col_widths.iter().sum();
    let table_x_offset = match table.style.alignment.as_deref() {
        Some("center") => x_offset + (available_width - table_total_width).max(0.0) / 2.0,
        Some("right") => x_offset + (available_width - table_total_width).max(0.0),
        _ => x_offset + table.style.indent.unwrap_or(0.0) as f64,
    };

    // Border style from document
    let border_color = table.style.border_color.as_deref()
        .and_then(parse_hex_color)
        .unwrap_or(Color::Gray(0.0)); // Word default: black borders
    let border_width = table.style.border_width.map(|w| w as f64).unwrap_or(0.5);

    let row_min_height = 14.0;

    for (_row_idx, row) in table.rows.iter().enumerate() {
        // Compute cell widths from grid columns + gridSpan
        // Handle rows where total gridSpan < grid_col count (vMerge or structural columns)
        let mut cell_widths = Vec::new();
        let total_span: usize = row.cells.iter().map(|c| c.grid_span.max(1) as usize).sum();
        let num_grid = grid_col_widths.len();
        // If row spans fewer columns than grid, skip leading columns
        let start_grid = if total_span < num_grid { num_grid - total_span } else { 0 };
        let mut grid_idx = start_grid;
        for cell in &row.cells {
            let span = cell.grid_span.max(1) as usize;
            let end = (grid_idx + span).min(num_grid);
            let cw: f64 = if grid_idx < num_grid {
                grid_col_widths[grid_idx..end].iter().sum()
            } else {
                // Fallback: use cell's own width if available
                cell.width.unwrap_or(0.0) as f64
            };
            // If computed grid width is tiny but cell has a specified width, use cell width
            let cw = if cw < 5.0 {
                cell.width.map(|w| w as f64).unwrap_or(cw).max(cw)
            } else {
                cw
            };
            cell_widths.push(cw);
            grid_idx = end;
            if grid_idx >= num_grid { break; }
        }

        // Calculate row height based on content (skip vMerge continue cells)
        let is_exact_height = row.height_rule.as_deref() == Some("exact");
        let mut row_height = row.height.map(|h| h as f64).unwrap_or(row_min_height);
        if !is_exact_height {
            // atLeast (default): max of specified height and content
            for (ci, cell) in row.cells.iter().enumerate() {
                if ci >= cell_widths.len() { break; }
                if cell.v_merge.as_deref() == Some("continue") { continue; }
                let content_w = (cell_widths[ci] - cell_margin_l - cell_margin_r).max(1.0);
                let cell_h = estimate_cell_height(cell, content_w, registry, doc_default_font, &builder.font_widths);
                let needed = cell_h + cell_margin_t + cell_margin_b;
                if needed > row_height {
                    row_height = needed;
                }
            }
        }
        // exact: use specified height regardless of content

        if builder.needs_page_break(row_height + 2.0) {
            builder.new_page();
        }

        let row_top = builder.cursor_y;
        let row_bottom = row_top + row_height;

        // Draw cells — offset by any skipped leading grid columns
        let leading_width: f64 = grid_col_widths[..start_grid].iter().sum();
        let mut cell_x = table_x_offset + leading_width;
        for (ci, cell) in row.cells.iter().enumerate() {
            if ci >= cell_widths.len() { break; }
            let cw = cell_widths[ci];
            let is_vmerge_continue = cell.v_merge.as_deref() == Some("continue");

            // Cell shading from document
            if let Some(ref shade) = cell.shading {
                if let Some(bg_color) = parse_hex_color(shade) {
                    builder.add_rect_fill(cell_x, row_top, cw, row_height, bg_color);
                }
            }

            // Cell borders: use cell-specific borders if available, else table defaults
            if table.style.border || cell.borders.is_some() {
                let draw_cell_border = |builder: &mut PdfBuilder, bd: Option<&oxidocs_core::ir::BorderDef>, x1: f64, y1: f64, x2: f64, y2: f64, default_draw: bool| {
                    match bd {
                        Some(bd) if bd.style == "nil" || bd.style == "none" => {},
                        Some(bd) => {
                            let bw = if bd.width > 0.0 { bd.width as f64 } else { border_width };
                            let bc = bd.color.as_deref().and_then(parse_hex_color).unwrap_or(border_color);
                            builder.add_line(x1, y1, x2, y2, bw, bc);
                        },
                        None if default_draw => {
                            builder.add_line(x1, y1, x2, y2, border_width, border_color);
                        },
                        _ => {},
                    }
                };
                let cb = cell.borders.as_ref();
                let has_table_border = table.style.border;
                draw_cell_border(builder, cb.and_then(|b| b.top.as_ref()), cell_x, row_top, cell_x + cw, row_top, has_table_border);
                draw_cell_border(builder, cb.and_then(|b| b.bottom.as_ref()), cell_x, row_bottom, cell_x + cw, row_bottom, has_table_border);
                draw_cell_border(builder, cb.and_then(|b| b.left.as_ref()), cell_x, row_top, cell_x, row_bottom, has_table_border);
                draw_cell_border(builder, cb.and_then(|b| b.right.as_ref()), cell_x + cw, row_top, cell_x + cw, row_bottom, has_table_border);
            }

            // Skip content rendering for vertically merged continuation cells
            if is_vmerge_continue {
                cell_x += cw;
                continue;
            }

            // Cell text — use cell-specific margins if available, else table defaults
            let cm_l = cell.margins.as_ref().and_then(|m| m.left).unwrap_or(cell_margin_l as f32) as f64;
            let cm_r = cell.margins.as_ref().and_then(|m| m.right).unwrap_or(cell_margin_r as f32) as f64;
            let cm_t = cell.margins.as_ref().and_then(|m| m.top).unwrap_or(cell_margin_t as f32) as f64;
            let cell_content_x = cell_x + cm_l;
            let cell_content_width = (cw - cm_l - cm_r).max(1.0);
            let mut cell_text_y = row_top + cm_t;

            for block in &cell.blocks {
                if let Block::Paragraph(para) = block {
                    let (font_size, cell_font, bold) = para_font_props(para);
                    let line_height = compute_line_height(
                        registry, cell_font, font_size as f32,
                        para.style.line_spacing, para.style.line_spacing_rule.as_deref(),
                        doc_default_font, 0.0, // no grid snap in table cells
                    );

                    cell_text_y += para.style.space_before.unwrap_or(0.0) as f64;

                    // Clip content that overflows exact row height
                    if is_exact_height && cell_text_y >= row_bottom {
                        break;
                    }

                    let text: String = para.runs.iter().map(|r| r.text.as_str()).collect();
                    if text.trim().is_empty() {
                        cell_text_y += line_height;
                    } else {
                        // Use per-run wrapping with ASCII/CJK font splitting
                        let cell_default_bold = para.style.default_run_style.as_ref().map(|rs| rs.bold).unwrap_or(false);
                        let cell_default_ff = para.style.default_run_style.as_ref().and_then(|rs| rs.font_family.as_deref());
                        let cell_lines = wrap_runs_into_lines(&para.runs, font_size, cell_content_width, 0.0, &builder.font_widths, cell_default_bold, cell_default_ff);
                        let text_color = para_color(para);

                        for line_runs in &cell_lines {
                            let line_width: f64 = line_runs.iter()
                                .map(|lr| {
                                    let base_w = builder.text_width(&lr.text, lr.font_size);
                                    base_w + lr.char_spacing * lr.text.chars().count() as f64
                                })
                                .sum();
                            let align_off = match para.alignment {
                                Alignment::Center => (cell_content_width - line_width) / 2.0,
                                Alignment::Right => cell_content_width - line_width,
                                _ => 0.0,
                            };

                            // Advance cursor FIRST (baseline at cell_text_y + line_height)
                            cell_text_y += line_height;
                            // Clip text that overflows exact row height
                            if is_exact_height && cell_text_y > row_bottom {
                                break;
                            }
                            let mut run_x = cell_content_x + align_off.max(0.0);
                            for lr in line_runs {
                                if !lr.text.is_empty() {
                                    if lr.char_spacing.abs() > 0.01 {
                                        for ch in lr.text.chars() {
                                            let ch_str = ch.to_string();
                                            builder.add_text(
                                                run_x, cell_text_y,
                                                ch_str.clone(), lr.font_name.clone(), lr.font_size, text_color,
                                            );
                                            run_x += builder.text_width(&ch_str, lr.font_size) + lr.char_spacing;
                                        }
                                    } else {
                                        builder.add_text(
                                            run_x, cell_text_y,
                                            lr.text.clone(), lr.font_name.clone(), lr.font_size, text_color,
                                        );
                                        run_x += builder.text_width(&lr.text, lr.font_size);
                                    }
                                }
                            }
                        }
                    }
                    cell_text_y += para.style.space_after.unwrap_or(0.0) as f64;
                }
            }

            cell_x += cw;
        }

        builder.cursor_y = row_bottom;
    }
}

fn estimate_cell_height(cell: &oxidocs_core::ir::TableCell, content_width: f64, registry: &FontMetricsRegistry, doc_default_font: Option<&str>, font_widths: &HashMap<u32, u16>) -> f64 {
    let mut h = 0.0;
    for block in &cell.blocks {
        if let Block::Paragraph(para) = block {
            let (font_size, cell_font, default_bold) = para_font_props(para);
            let line_height = compute_line_height(
                registry, cell_font, font_size as f32,
                para.style.line_spacing, para.style.line_spacing_rule.as_deref(),
                doc_default_font, 0.0, // no grid snap in table cells
            );
            h += para.style.space_before.unwrap_or(0.0) as f64;
            let text: String = para.runs.iter().map(|r| r.text.as_str()).collect();
            if text.trim().is_empty() {
                h += line_height;
            } else {
                // Use run-aware wrapping with actual font widths (same as rendering pass)
                let default_ff = para.style.default_run_style.as_ref().and_then(|rs| rs.font_family.as_deref());
                let cell_lines = wrap_runs_into_lines(&para.runs, font_size, content_width, 0.0, font_widths, default_bold, default_ff);
                h += cell_lines.len() as f64 * line_height;
            }
            h += para.style.space_after.unwrap_or(0.0) as f64;
        }
    }
    h.max(12.0)
}

// --- Font Embedding ---

/// System font paths to search for CJK fonts (Regular)
const CJK_FONT_PATHS: &[&str] = &[
    // Prefer MS Gothic (most common in Japanese government documents)
    "C:\\Windows\\Fonts\\msgothic.ttc",
    "C:\\Windows\\Fonts\\YuGothR.ttc",
    "C:\\Windows\\Fonts\\meiryo.ttc",
    "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc",
    "/System/Library/Fonts/ヒラギノ角ゴシック W3.ttc",
    "/System/Library/Fonts/HiraginoSans-W3.ttc",
];

/// System font paths to search for CJK fonts (Bold)
const CJK_FONT_PATHS_BOLD: &[&str] = &[
    "C:\\Windows\\Fonts\\msgothic.ttc",
    "C:\\Windows\\Fonts\\YuGothB.ttc",
    "C:\\Windows\\Fonts\\meiryo.ttc",
    "/usr/share/fonts/opentype/noto/NotoSansCJK-Bold.ttc",
    "/System/Library/Fonts/ヒラギノ角ゴシック W6.ttc",
    "/System/Library/Fonts/HiraginoSans-W6.ttc",
];

/// Path to Python3 for font subsetting (requires fonttools)
const PYTHON3: &str = "python3";

/// Faces asked, in order, for a character no face of the document has: the
/// check boxes, ≒ and Wingdings bullets of a Japanese form drawn in a stand-in
/// Gothic that lacks them. Read off the machine, never shipped: DejaVu and the
/// Noto symbol faces are in the Debian images a server runs (Gotenberg ships
/// both), Segoe UI Symbol on Windows.
const FALLBACK_FACES: &[&str] = &[
    "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
    "/usr/share/fonts/truetype/noto/NotoSansSymbols2-Regular.ttf",
    "/usr/share/fonts/truetype/noto/NotoSansSymbols-Regular.ttf",
    "/usr/share/fonts/truetype/noto/NotoSansMath-Regular.ttf",
    "C:\\Windows\\Fonts\\seguisym.ttf",
];

/// PDF font name of the `n`th fallback face that was embedded.
fn fallback_face_name(n: usize) -> String {
    format!("OxiFallback-{n}")
}

/// Whether `font_name` is an East Asian face, which then draws every
/// character of its run rather than handing the Latin ones to another face.
fn is_cjk_pdf_font(font_name: &str) -> bool {
    ["Gothic", "Mincho", "ゴシック", "明朝", "游", "ＭＳ", "メイリオ", "Meiryo"]
        .iter()
        .any(|k| font_name.contains(k))
}

/// `text` with a symbol font's codes turned into the Unicode symbols they
/// draw, for each code the face embedded under the font's name has no glyph
/// for. The real Wingdings maps its codes (at U+F0xx, however the run spells
/// them) and keeps drawing them; the face a machine without it substitutes
/// under the same name does not.
fn symbol_text(family: Option<&str>, font_name: &str, text: &str, fonts: &HashMap<String, EmbeddedFont>) -> String {
    let Some(family) = family else { return text.to_string() };
    let face = fonts.get(font_name);
    text.chars()
        .map(|c| match oxidocs_fonts::symbol_font_to_unicode(family, c) {
            Some(u) if face.map_or(true, |f| !f.unicode_to_gid.contains_key(&(0xF000 | (c as u32 & 0xFF)))) => u,
            _ => c,
        })
        .collect()
}

/// The face a character of a run in `font_name` is drawn with, before any
/// fallback: the run's face, or the East Asian face for what a Latin face
/// lacks.
fn primary_face_for(c: char, font_name: &str, bold: bool, fonts: &HashMap<String, EmbeddedFont>) -> String {
    // An East Asian face draws its whole run.
    if is_cjk_pdf_font(font_name) {
        return font_name.to_string();
    }
    // Only divert to a font that was actually loaded. The decision to load one
    // is made from the document's text; this one is made per character, and
    // when the two disagreed a bullet missing from the Latin subset sent a run
    // to OxiCJK in a document with no East Asian text, naming a font the PDF
    // never carried.
    // A tab carries no glyph and is in no font's cmap, so "the face does not
    // cover it" is true of every face -- which used to hand a plain English
    // document to the East Asian font on the strength of one indent.
    if !fonts.contains_key("OxiCJK-Regular") || (c as u32) < 0x20 {
        return font_name.to_string();
    }
    // Route to the CJK font only when the resolved face has no glyph for the
    // character. The old test was "codepoint > 0x7F", which sent curly quotes,
    // dashes and accented Latin into a Japanese Gothic font. With no face to
    // ask, only real East Asian text is worth diverting.
    let needs_cjk = match fonts.get(font_name) {
        Some(face) => !face.unicode_to_gid.contains_key(&(c as u32)),
        None => is_cjk_char(c),
    };
    if !needs_cjk {
        font_name.to_string()
    } else if bold {
        "OxiCJK-Bold".to_string()
    } else {
        "OxiCJK-Regular".to_string()
    }
}

/// Whether the face `name` is embedded and has no glyph for `c` -- the case a
/// fallback face exists for. Spaces and control characters draw nothing and
/// need no glyph.
fn lacks_glyph(name: &str, c: char, fonts: &HashMap<String, EmbeddedFont>) -> bool {
    !c.is_whitespace()
        && (c as u32) >= 0x20
        && fonts.get(name).is_some_and(|face| !face.unicode_to_gid.contains_key(&(c as u32)))
}

/// The face a character is drawn with: its primary face, or the first
/// embedded fallback face that has the glyph the primary one lacks.
fn pdf_face_for(c: char, font_name: &str, bold: bool, fonts: &HashMap<String, EmbeddedFont>) -> String {
    let primary = primary_face_for(c, font_name, bold, fonts);
    if lacks_glyph(&primary, c, fonts) {
        for n in 0.. {
            let Some(face) = fonts.get(&fallback_face_name(n)) else { break };
            if face.unicode_to_gid.contains_key(&(c as u32)) {
                return fallback_face_name(n);
            }
        }
    }
    primary
}

/// Embed, from `FALLBACK_FACES`, the glyphs of every character whose primary
/// face lacks it, so it is drawn instead of left as .notdef.
fn embed_fallback_faces(layout: &oxidocs_core::layout::LayoutResult, fonts: &mut HashMap<String, EmbeddedFont>) {
    let mut missing: std::collections::BTreeSet<char> = std::collections::BTreeSet::new();
    for page in &layout.pages {
        for elem in &page.elements {
            if let oxidocs_core::layout::LayoutContent::Text { text, font_family, bold, italic, .. } = &elem.content {
                let font_name = resolve_font_styled(font_family.as_deref(), *bold, *italic);
                for c in symbol_text(font_family.as_deref(), &font_name, text, fonts).chars() {
                    if lacks_glyph(&primary_face_for(c, &font_name, *bold, fonts), c, fonts) {
                        missing.insert(c);
                    }
                }
            }
        }
    }
    let mut n = 0;
    for path in FALLBACK_FACES {
        if missing.is_empty() {
            break;
        }
        if !std::path::Path::new(path).exists() {
            continue;
        }
        let Some(face) = subset_font_file(path, &missing, &format!("fallback{n}")) else { continue };
        let covered: Vec<char> = missing.iter().copied().filter(|c| face.unicode_to_gid.contains_key(&(*c as u32))).collect();
        if covered.is_empty() {
            continue;
        }
        note!("Embedding fallback face {} for {} character(s) no document face has", path, covered.len());
        for c in covered {
            missing.remove(&c);
        }
        fonts.insert(fallback_face_name(n), face);
        n += 1;
    }
    if !missing.is_empty() {
        let list: String = missing.iter().collect();
        eprintln!("Warning: no face on this machine has {list:?} - drawn as missing glyphs");
    }
}

/// True for characters that actually need an East Asian face.
///
/// The test used to be `c > 0x7F`, which is true of a curly quote, an en dash
/// and every accented Latin letter. One smart quote in an English document was
/// enough to decide it "needs CJK" and pull a whole East Asian font into the
/// PDF — 15MB of Noto on Linux, for a document with no East Asian text in it.
fn is_cjk_char(c: char) -> bool {
    matches!(c as u32,
        0x2E80..=0x2FFF   // radicals and Kangxi
        | 0x3000..=0x303F // CJK symbols and punctuation
        | 0x3040..=0x30FF // hiragana, katakana
        | 0x3100..=0x312F // bopomofo
        | 0x3130..=0x318F // hangul compatibility jamo
        | 0x3190..=0x319F // kanbun
        | 0x31F0..=0x31FF // katakana phonetic extensions
        | 0x3400..=0x4DBF // unified ideographs extension A
        | 0x4E00..=0x9FFF // unified ideographs
        | 0xA960..=0xA97F // hangul jamo extended-A
        | 0xAC00..=0xD7AF // hangul syllables
        | 0xF900..=0xFAFF // compatibility ideographs
        | 0xFE30..=0xFE4F // CJK compatibility forms
        | 0xFF00..=0xFFEF // halfwidth and fullwidth forms
        | 0x1100..=0x11FF // hangul jamo
        | 0x20000..=0x3FFFF // ideograph extensions B and later
    )
}

/// True for the East Asian families the CJK embedding path owns. The generic
/// face pass leaves these alone so the existing per-family CJK handling (and
/// the emitter's ASCII/CJK segment split, which keys off the returned name)
/// keeps working unchanged.
fn is_cjk_family(base: &str) -> bool {
    // A family name written in Japanese is East Asian whatever else it says.
    if base.chars().any(is_cjk_char) {
        return true;
    }
    // Romanized names have to be matched on words, not substrings. "Gothic"
    // alone was the test, and it is a Latin word: Century Gothic, Franklin
    // Gothic and News Gothic are all Latin families that were being handed to
    // the East Asian path and drawn with MS Gothic.
    let lower = base.to_lowercase();
    const EAST_ASIAN: &[&str] = &[
        "ms gothic", "ms pgothic", "ms mincho", "ms pmincho",
        "yu gothic", "yu mincho", "meiryo", "hiragino", "noto sans cjk",
        "noto serif cjk", "source han", "pingfang", "heiti", "songti",
        "kaiti", "fangsong", "simsun", "simhei", "nsimsun", "batang",
        "gulim", "dotum", "malgun", "nanum",
    ];
    EAST_ASIAN.iter().any(|name| lower.starts_with(name))
}

/// Resolve every (family, weight, slope) the laid-out document asks for to a
/// real font file, subset it to the characters that face is used with, and
/// embed it under a stable key.
///
/// Order: the family itself, then the bundled metric-compatible substitute,
/// then nothing -- in which case `resolve_font` falls back to the base-14
/// mapping and the viewer substitutes. That last case is reported on stderr,
/// because a PDF built from a substituted face looks fine and is wrong, and
/// silence is what made it invisible before.
fn embed_document_faces(
    layout: &oxidocs_core::layout::LayoutResult,
    fonts: &mut HashMap<String, EmbeddedFont>,
) {
    use std::collections::{BTreeMap, BTreeSet};

    // A plan left by this thread's previous conversion must not survive into
    // this one, even when this document embeds nothing.
    FONT_PLAN.with(|p| *p.borrow_mut() = FontPlan::default());
    let mut wanted: BTreeMap<(String, bool, bool), BTreeSet<char>> = BTreeMap::new();
    for page in &layout.pages {
        for elem in &page.elements {
            if let oxidocs_core::layout::LayoutContent::Text {
                text, font_family, bold, italic, ..
            } = &elem.content
            {
                let family = font_family
                    .clone()
                    .unwrap_or_else(|| "Calibri".to_string());
                if is_cjk_family(&family) {
                    continue;
                }
                let set = wanted.entry((family, *bold, *italic)).or_default();
                for ch in text.chars() {
                    set.insert(ch);
                }
            }
            // OXI_PDF_EMF: the faces a metafile sets its text in.
            if let oxidocs_core::layout::LayoutContent::Image { data, .. } = &elem.content {
                for (face, bold, italic, text) in emf::text_runs(data).into_iter().chain(emf::wmf_text_runs(data)) {
                    if is_cjk_family(&face) {
                        continue;
                    }
                    wanted.entry((face, bold, italic)).or_default().extend(text.chars());
                }
            }
        }
    }
    if wanted.is_empty() {
        return;
    }

    let index = font_index();
    let mut plan = FontPlan::default();
    let mut missing: Vec<String> = Vec::new();

    for ((family, bold, italic), chars) in wanted {
        // A face asked for by an empty run draws nothing. Subsetting it to
        // zero glyphs drops the cmap and the subsetter dies on the way out,
        // so there is nothing to embed and nothing to warn about.
        if chars.is_empty() {
            continue;
        }
        let key = font_key(&family, bold, italic);
        if fonts.contains_key(&key) {
            plan.keys
                .insert((family.trim().to_lowercase(), bold, italic), key);
            continue;
        }
        // The family itself, then its stand-in, then the same pair again with
        // a weight or slope word taken off the end — `Calibri Light` is a
        // Calibri a document wrote the long way, and it is the only handle we
        // have on the face it means.
        let lookup = |name: &str, bold: bool, italic: bool| {
            index
                .find(name, bold, italic)
                .map(|face| (face, None))
                .or_else(|| {
                    substitute_for(name)
                        .and_then(|s| index.find(s, bold, italic))
                        .map(|face| (face, substitute_for(name)))
                })
        };
        let (face, substituted) = match lookup(&family, bold, italic) {
            Some((face, sub)) => (Some(face), sub),
            None => match split_style_suffix(&family)
                .and_then(|(base, b, i)| lookup(&base, bold || b, italic || i))
            {
                Some((face, sub)) => (Some(face), sub),
                None => (None, None),
            },
        };
        let Some(face) = face else {
            missing.push(family.clone());
            continue;
        };
        let path = face.path.to_string_lossy().to_string();
        // The external subsetter addresses files, not collection members, so a
        // face past the first in its .ttc goes out whole instead. macOS ships
        // most of its families that way — Times.ttc holds roman, bold, italic
        // and bold-italic as members 0..3 — so refusing them would leave every
        // emphasised run on that platform drawn by a viewer substitute.
        let embedded = if face.index == 0 {
            subset_font_file(&path, &chars, &key)
        } else {
            embed_collection_member(&path, face.index, &key)
        };
        match embedded {
            Some(embedded) => {
                if let Some(sub) = substituted {
                    note!(
                        "Substituting {sub} for {family} (not installed; advance widths match on ASCII)"
                    );
                }
                fonts.insert(key.clone(), embedded);
                plan.keys
                    .insert((family.trim().to_lowercase(), bold, italic), key);
            }
            None => missing.push(family.clone()),
        }
    }

    if !missing.is_empty() {
        missing.sort();
        missing.dedup();
        eprintln!(
            "Warning: no font file for {} - these are drawn with a viewer substitute and will not match Word",
            missing.join(", ")
        );
    }

    FONT_PLAN.with(|p| *p.borrow_mut() = plan);
}

/// Build embedded fonts map from system fonts.
/// Detects CJK font usage in the document and embeds the font data.
fn build_embedded_fonts(
    doc: &oxidocs_core::Document,
    layout: &oxidocs_core::layout::LayoutResult,
) -> HashMap<String, EmbeddedFont> {
    let mut fonts = HashMap::new();

    // Every face the laid-out document actually asks for, with the characters
    // it asks for. Read from the layout rather than the IR because bold and
    // italic are resolved there -- the IR carries the level, not the answer.
    embed_document_faces(layout, &mut fonts);

    // Collect all font names used in the document that likely need CJK
    let mut needs_cjk_font = false;
    let mut font_names_used: Vec<String> = Vec::new();

    for page in &doc.pages {
        for block in &page.blocks {
            collect_cjk_fonts_from_block(block, &mut needs_cjk_font, &mut font_names_used);
        }
    }

    // Also collect document-level default fonts (from rPrDefault / styles)
    if let Some(drs) = doc.styles.effective_doc_default_run_style() {
        if let Some(ref ff) = drs.font_family_east_asia {
            if !font_names_used.contains(ff) {
                font_names_used.push(ff.clone());
            }
        }
        if let Some(ref ff) = drs.font_family {
            if !font_names_used.contains(ff) {
                font_names_used.push(ff.clone());
            }
        }
    }

    // OXI_PDF_EMF_CJK (2026-11-12, default ON, opt-out _DISABLE): the text a
    // metafile draws, as (face, bold, italic, text). Its East Asian faces and
    // characters were never embedded -- the families and the subset came from
    // the layout's text runs only -- so a WMF set in ＭＳ Ｐゴシック went out
    // under a bare MS-PGothic no viewer can draw: blind-K
    // ja/creative__9bf674f6, a school plan whose right column is one WMF.
    let metafile_runs: Vec<(String, bool, bool, String)> = if std::env::var_os("OXI_PDF_EMF_CJK_DISABLE").is_some() {
        Vec::new()
    } else {
        layout
            .pages
            .iter()
            .flat_map(|page| &page.elements)
            .flat_map(|elem| match &elem.content {
                oxidocs_core::layout::LayoutContent::Image { data, .. } => {
                    emf::text_runs(data).into_iter().chain(emf::wmf_text_runs(data)).collect()
                }
                _ => Vec::new(),
            })
            .collect()
    };

    // Every character the pages draw. The IR's body blocks miss whatever the
    // layout produces or pulls in itself -- page numbers and other field
    // results, list labels, headers, footers, footnotes, shape text -- and a
    // character missing from the subset is drawn as .notdef.
    let mut drawn_chars: std::collections::BTreeSet<char> = std::collections::BTreeSet::new();
    for (_, _, _, text) in &metafile_runs {
        drawn_chars.extend(text.chars());
    }
    for page in &layout.pages {
        for elem in &page.elements {
            if let oxidocs_core::layout::LayoutContent::Text { text, font_family, bold, italic, .. } = &elem.content {
                drawn_chars.extend(text.chars());
                // and the Unicode a symbol font's codes are drawn as without it
                let font_name = resolve_font_styled(font_family.as_deref(), *bold, *italic);
                drawn_chars.extend(symbol_text(font_family.as_deref(), &font_name, text, &fonts).chars());
            }
        }
    }

    // Embed Latin fonts (Calibri + Cambria) regardless of CJK usage
    let mut used_chars_latin: std::collections::BTreeSet<char> = drawn_chars.clone();
    for page in &doc.pages {
        for block in &page.blocks {
            collect_chars_from_block(block, &mut used_chars_latin);
        }
    }
    embed_latin_fonts(&used_chars_latin, &mut fonts);
    embed_cambria_fonts(&used_chars_latin, &mut fonts);

    // The body blocks are not the only place East Asian text lives: a form
    // drawn in text boxes, a header, a shape's label. Deciding from them alone
    // left such a document with no East Asian face at all, and its text went
    // out under a bare font name no viewer can draw it with.
    // An East Asian family drawing only Latin text -- MS Gothic for the digits
    // of an English form -- needs its face as much: its runs are named after
    // it whatever they hold.
    let east_asian_runs = layout.pages.iter().flat_map(|page| &page.elements).any(|elem| {
        matches!(&elem.content, oxidocs_core::layout::LayoutContent::Text { text, font_family: Some(family), .. }
            if is_cjk_family(family) && !text.trim().is_empty())
    }) || metafile_runs.iter().any(|(face, _, _, text)| is_cjk_family(face) && !text.trim().is_empty());
    if east_asian_runs || drawn_chars.iter().any(|c| is_cjk_char(*c)) {
        needs_cjk_font = true;
    }
    if !needs_cjk_font {
        return fonts;
    }

    // Collect all unique characters used in the document (for font subsetting)
    let mut used_chars: std::collections::BTreeSet<char> = drawn_chars;
    for page in &doc.pages {
        for block in &page.blocks {
            collect_chars_from_block(block, &mut used_chars);
        }
        // Also collect characters from TextBoxes
        for text_box in &page.text_boxes {
            for block in &text_box.blocks {
                collect_chars_from_block(block, &mut used_chars);
            }
        }
    }

    // The East Asian families the text is actually drawn in. The IR alone
    // misses a family that only a style names -- ＭＳ 明朝 on Normal's
    // eastAsia slot is the common case -- and the layout has already resolved
    // every run to the family it draws with.
    for page in &layout.pages {
        for elem in &page.elements {
            if let oxidocs_core::layout::LayoutContent::Text { font_family: Some(family), .. } = &elem.content {
                if is_cjk_family(family) && !font_names_used.contains(family) {
                    font_names_used.push(family.clone());
                }
            }
        }
    }
    for (face, _, _, _) in &metafile_runs {
        if is_cjk_family(face) && !font_names_used.contains(face) {
            font_names_used.push(face.clone());
        }
    }

    if cjk_bold_face_enabled() {
        let mut bold_families: Vec<String> = Vec::new();
        for page in &layout.pages {
            for elem in &page.elements {
                if let oxidocs_core::layout::LayoutContent::Text { font_family: Some(family), bold: true, .. } = &elem.content {
                    if is_cjk_family(family) && !bold_families.contains(family) {
                        bold_families.push(family.clone());
                    }
                }
            }
        }
        for (face, bold, _, _) in &metafile_runs {
            if *bold && is_cjk_family(face) && !bold_families.contains(face) {
                bold_families.push(face.clone());
            }
        }
        for name in &bold_families {
            let key = resolve_font_styled(Some(name), true, false);
            if fonts.contains_key(&key) || !key.ends_with("-Bold") || key == "OxiCJK-Bold" {
                continue;
            }
            if let Some((path, face)) = face_file_cached(name, true, false) {
                if let Ok(result) = subset_font_native_face(&path.to_string_lossy(), face, &used_chars) {
                    note!("Embedding CJK bold font '{}' ({}) from {} face {} ({} bytes)", key, name, path.display(), face, result.data.len());
                    fonts.insert(key, result);
                }
            }
        }
    }

    // Per-family CJK font embedding: the family's own file where the machine
    // has it, then the shipped stand-in of the same kind -- Mincho for a
    // Mincho family, Gothic otherwise. The stand-in is looked for where every
    // other shipped face is, so a Mincho document is never handed to whatever
    // sans face the system happens to carry.
    let system_files: &[(&str, &str)] = &[
        ("MS-Gothic", "C:\\Windows\\Fonts\\msgothic.ttc"),
        ("MS-PGothic", "C:\\Windows\\Fonts\\msgothic.ttc"),
        ("MS-Mincho", "C:\\Windows\\Fonts\\msmincho.ttc"),
        ("MS-PMincho", "C:\\Windows\\Fonts\\msmincho.ttc"),
    ];
    for name in &font_names_used {
        if !is_cjk_family(name) {
            continue;
        }
        let resolved = resolve_font(Some(name), false);
        if fonts.contains_key(&resolved) {
            continue;
        }
        // OXI_PDF_FACE_INDEX (2026-11-08, default ON, opt-out _DISABLE): embed the face
        // the document names from the machine's own file AND collection index, found
        // the way the layout finds it (font::runtime). The table below always took face
        // 0 of a .ttc, so ＭＳ Ｐゴシック came out as ＭＳ ゴシック (msgothic.ttc face 0,
        // monospaced) and 游明朝 / メイリオ fell to the stand-in: blind-J ja/reports__
        // 77526bdb's proportional body ran full-width past its cells in the PDF while
        // the layout (and the renderer) set it at the proportional advances.
        if std::env::var_os("OXI_PDF_FACE_INDEX_DISABLE").is_none() {
            if let Some((path, face)) = face_file_cached(name, false, false) {
                let label = resolved.replace(['-', ' '], "_");
                if let Ok(result) = subset_font_native_face(&path.to_string_lossy(), face, &used_chars) {
                    note!("Embedding CJK font '{}' ({}) from {} face {} ({} bytes)", resolved, name, path.display(), face, result.data.len());
                    fonts.insert(resolved, result);
                    continue;
                }
            }
        }
        let system = system_files
            .iter()
            .find(|(family, _)| *family == resolved)
            .map(|(_, path)| std::path::PathBuf::from(path))
            .filter(|path| path.exists());
        let found_path = system.or_else(|| bundled_face_file(oxidocs_fonts::cjk_stand_in(name).1));
        let label = resolved.replace(['-', ' '], "_");
        if let Some(path) = found_path {
            if let Some(result) = subset_font_file(&path.to_string_lossy(), &used_chars, &label) {
                note!("Embedding CJK font '{}' ({}) from {} ({} bytes)", resolved, name, path.display(), result.data.len());
                fonts.insert(resolved, result);
            }
        }
    }

    // Try to find, subset, and load a default CJK font from the system (for OxiCJK-Regular)
    let font_result = match find_and_subset_cjk_font(&used_chars) {
        Some(result) => result,
        None => {
            eprintln!("Warning: No CJK system font found. PDF may not render Japanese text correctly.");
            return fonts;
        }
    };

    note!("Embedding CJK font ({} bytes, {} glyph mappings)",
        font_result.data.len(), font_result.unicode_to_gid.len());

    // Map remaining CJK font names (not already per-family embedded) to default CJK font
    for name in &font_names_used {
        let resolved = resolve_font(Some(name), false);
        if !fonts.contains_key(&resolved) {
            fonts.insert(resolved.clone(), font_result.clone());
        }
        let resolved_bold = resolve_font(Some(name), true);
        if !fonts.contains_key(&resolved_bold) {
            fonts.insert(resolved_bold, font_result.clone());
        }
    }

    // Map the dedicated CJK font names used for non-ASCII text spans
    fonts.insert("OxiCJK-Regular".to_string(), font_result.clone());

    // Try to embed a Bold CJK font separately
    if let Some(bold_result) = find_and_subset_cjk_font_bold(&used_chars) {
        note!("Embedding CJK Bold font ({} bytes, {} glyph mappings)",
            bold_result.data.len(), bold_result.unicode_to_gid.len());
        fonts.insert("OxiCJK-Bold".to_string(), bold_result);
    } else {
        // Fallback: use regular font for bold too
        fonts.insert("OxiCJK-Bold".to_string(), font_result.clone());
    }

    // Latin fonts already embedded above (before CJK check)

    fonts
}

/// Embed Cambria Regular and Bold for documents using Cambria/Century.
fn embed_cambria_fonts(used_chars: &std::collections::BTreeSet<char>, fonts: &mut HashMap<String, EmbeddedFont>) {
    let regular = match resolve_face_path("Cambria", false, false) {
        Some(path) => path,
        None => return,
    };
    let bold = match resolve_face_path("Cambria", true, false) {
        Some(path) => path,
        None => regular.clone(),
    };
    let (regular_path, bold_path) = (regular.as_str(), bold.as_str());
    let latin_chars: std::collections::BTreeSet<char> = used_chars.iter()
        .filter(|c| (**c as u32) < 0x2000)
        .copied()
        .collect();
    if let Some(result) = subset_font_file(regular_path, &latin_chars, "cambria") {
        note!("Embedding Cambria font ({} bytes, {} glyphs)", result.data.len(), result.unicode_to_gid.len());
        fonts.insert("OxiCambria-Regular".to_string(), result);
    }
    if let Some(result) = subset_font_file(bold_path, &latin_chars, "cambria-bold") {
        note!("Embedding Cambria Bold font ({} bytes, {} glyphs)", result.data.len(), result.unicode_to_gid.len());
        fonts.insert("OxiCambria-Bold".to_string(), result);
    }
}

/// Embed Calibri Regular and Bold for Latin text rendering.
fn embed_latin_fonts(used_chars: &std::collections::BTreeSet<char>, fonts: &mut HashMap<String, EmbeddedFont>) {
    // Fallback faces for the `OxiLatin-*` names `resolve_font` still returns
    // when a document's own family could not be embedded. Resolved through the
    // index like everything else, so a machine without Calibri gets Carlito
    // here too instead of falling off the end.
    let regular = match resolve_face_path("Calibri", false, false) {
        Some(path) => path,
        None => { eprintln!("Warning: No Latin font found"); return; }
    };
    let bold = match resolve_face_path("Calibri", true, false) {
        Some(path) => path,
        None => regular.clone(),
    };
    let (regular_path, bold_path) = (regular.as_str(), bold.as_str());

    // Filter to ASCII/Latin chars only (for smaller subset)
    let latin_chars: std::collections::BTreeSet<char> = used_chars.iter()
        .filter(|c| (**c as u32) < 0x2000)
        .copied()
        .collect();

    if let Some(result) = subset_font_file(regular_path, &latin_chars, "latin") {
        note!("Embedding Latin font ({} bytes, {} glyphs)", result.data.len(), result.unicode_to_gid.len());
        fonts.insert("OxiLatin-Regular".to_string(), result);
    }
    if let Some(result) = subset_font_file(bold_path, &latin_chars, "latin-bold") {
        note!("Embedding Latin Bold font ({} bytes, {} glyphs)", result.data.len(), result.unicode_to_gid.len());
        fonts.insert("OxiLatin-Bold".to_string(), result);
    }
}

/// Embed a font file whole, with no external help. This is the road out when
/// the subsetter cannot run: a machine with no `python3`, or no fontTools, or
/// a font it refuses. The PDF gets larger -- a CJK face is megabytes where its
/// subset is kilobytes -- but it is a correct PDF with real outlines and a
/// real width table, which a base-14 name is not. Being unable to find Python
/// must not decide how a document looks.
/// Remove what the subsetter wrote. Called on every way out of
/// `subset_font_file`, not just the successful one: `font.save()` happens
/// before the step that can fail, so a failing face still left three files
/// behind — 88 of them after thirty documents.
fn clear_scratch(widths_path: &str, cidmap_path: &str, subset_path: &str) {
    let _ = fs::remove_file(subset_path);
    let _ = fs::remove_file(cidmap_path);
    let _ = fs::remove_file(format!("{widths_path}.psname"));
    let _ = fs::remove_file(widths_path);
}

/// Embed one member of a font collection whole. Used where the family lives
/// past the first face of a `.ttc`, which the path-addressed subsetter cannot
/// name.
fn embed_collection_member(font_path: &str, face: u32, label: &str) -> Option<EmbeddedFont> {
    let data = fs::read(font_path).ok()?;
    let font = oxipdf_core::font_util::embedded_font_from_face(&data, face);
    if font.unicode_to_gid.is_empty() {
        return None;
    }
    note!(
        "Embedding {} whole from collection member {} ({} bytes)",
        label,
        face,
        font.data.len()
    );
    Some(font)
}

fn embed_whole_font(font_path: &str, label: &str) -> Option<EmbeddedFont> {
    let data = fs::read(font_path).ok()?;
    let font = oxipdf_core::font_util::embedded_font_from_ttf(&data);
    if font.unicode_to_gid.is_empty() {
        return None;
    }
    note!(
        "Embedding {} whole ({} bytes) — no subsetter available",
        label,
        font.data.len()
    );
    Some(font)
}

enum NativeSubset {
    /// The face draws none of the characters asked for.
    NoGlyph,
    /// The font is one the in-process subsetter cannot cut (CFF2, say).
    Failed(String),
}

/// Subset the first face of `font_path` to `chars` in-process. Glyphs are
/// renumbered from 1 in the order the characters come, and the new number is
/// the CID for TrueType and CFF alike, which is what the PDF writer encodes
/// with. The subset carries no cmap: `unicode_to_gid` is the mapping.
///
/// This replaced nine `python3 -c fontTools` runs per document -- three
/// quarters of a conversion's time, and a dependency a server image need
/// not have.
/// OXI_PDF_VERT_GLYPHS: an embedded face's vertical (GSUB `vert`) glyph for a character
/// is stored in `unicode_to_gid` under this bit plus the code point.
const VERT_GLYPH_KEY: u32 = 0x4000_0000;

/// Characters whose upright vertical form differs: CJK punctuation and brackets,
/// small kana, the long-vowel mark, full-width forms and dashes.
fn vert_candidate(c: char) -> bool {
    matches!(c as u32, 0x3000..=0x303F | 0xFF01..=0xFF60 | 0x2010..=0x2026 | 0x30FC | 0x2212)
        || "ぁぃぅぇぉっゃゅょゎゕゖァィゥェォッャュョヮヵヶㇰㇱㇲㇳㇴㇵㇶㇷㇸㇹㇺㇻㇼㇽㇾㇿ".contains(c)
}

fn subset_font_native(font_path: &str, chars: &std::collections::BTreeSet<char>) -> Result<EmbeddedFont, NativeSubset> {
    subset_font_native_face(font_path, 0, chars)
}

fn subset_font_native_face(font_path: &str, face: u32, chars: &std::collections::BTreeSet<char>) -> Result<EmbeddedFont, NativeSubset> {
    use oxipdf_core::font_util;
    let file = fs::read(font_path).map_err(|e| NativeSubset::Failed(e.to_string()))?;
    let lifted = font_util::extract_ttc_face(&file, face);
    let data: &[u8] = lifted.as_deref().unwrap_or(&file);
    // The best cmap, a symbol font's (3,0) table included, so a Wingdings
    // code finds its glyph.
    let cmap = font_util::parse_cmap_table(data);
    let widths = font_util::parse_ttf_widths(data);

    let mut remapper = subsetter::GlyphRemapper::new();
    let mut unicode_to_gid = HashMap::new();
    let mut cid_widths = HashMap::new();
    for &c in chars {
        let Some(&gid) = cmap.get(&(c as u32)) else { continue };
        if gid == 0 {
            continue;
        }
        let new = remapper.remap(gid);
        unicode_to_gid.insert(c as u32, new);
        if let Some(&w) = widths.get(&gid) {
            cid_widths.insert(new, w);
        }
    }
    if unicode_to_gid.is_empty() {
        return Err(NativeSubset::NoGlyph);
    }
    if std::env::var_os("OXI_PDF_VERT_GLYPHS_DISABLE").is_none() {
        if let Some(face) = rustybuzz::Face::from_slice(data, 0) {
            for &c in chars {
                if !vert_candidate(c) {
                    continue;
                }
                let Some(&gid) = cmap.get(&(c as u32)) else { continue };
                let mut buf = rustybuzz::UnicodeBuffer::new();
                buf.push_str(&c.to_string());
                buf.guess_segment_properties();
                buf.set_direction(rustybuzz::Direction::TopToBottom);
                let vert = rustybuzz::Feature::new(rustybuzz::ttf_parser::Tag::from_bytes(b"vert"), 1, ..);
                let shaped = rustybuzz::shape(&face, &[vert], buf);
                let info = shaped.glyph_infos();
                if info.len() != 1 || info[0].glyph_id == 0 || info[0].glyph_id == gid as u32 || info[0].glyph_id > u16::MAX as u32 {
                    continue;
                }
                let vg = info[0].glyph_id as u16;
                let new = remapper.remap(vg);
                unicode_to_gid.insert(VERT_GLYPH_KEY | c as u32, new);
                if let Some(&w) = widths.get(&vg) {
                    cid_widths.insert(new, w);
                }
            }
        }
    }
    let subset = subsetter::subset(data, 0, &remapper).map_err(|e| NativeSubset::Failed(e.to_string()))?;
    let ps_name = font_util::parse_ps_name(data);
    let (data, format) = if subset.starts_with(b"OTTO") || has_cff_table(&subset) {
        let cff = extract_cff_from_otf(&subset).ok_or_else(|| NativeSubset::Failed("no CFF table in the subset".into()))?;
        (cff, FontFormat::OpenTypeCff)
    } else {
        (subset, FontFormat::TrueType)
    };
    Ok(EmbeddedFont { ps_name, data, format, unicode_to_gid, cid_widths })
}

/// Subset a single font file and return EmbeddedFont. Falls back to embedding
/// the whole file when the external subsetter is unavailable or fails.
fn subset_font_file(font_path: &str, chars: &std::collections::BTreeSet<char>, label: &str) -> Option<EmbeddedFont> {
    // Nothing to draw, so nothing to embed. Subsetting to zero glyphs strips
    // the cmap and the subsetter dies reading it back; a document whose only
    // block is an empty table used to take every Latin face down that way.
    if chars.is_empty() {
        return None;
    }
    // OXI_NATIVE_SUBSET_DISABLE: go straight to fontTools, to compare.
    let native = if std::env::var_os("OXI_NATIVE_SUBSET_DISABLE").is_some() {
        Err(NativeSubset::Failed("disabled".into()))
    } else {
        subset_font_native(font_path, chars)
    };
    match native {
        Ok(font) => {
            note!("Font subset ({}): {} mappings, {} bytes", label, font.unicode_to_gid.len(), font.data.len());
            return Some(font);
        }
        Err(NativeSubset::NoGlyph) => return None,
        Err(NativeSubset::Failed(why)) => {
            note!("Font subset ({}) in-process failed ({}); trying fontTools", label, why);
        }
    }
    let char_string: String = chars.iter()
        .filter(|c| { let cp = **c as u32; !(0xD800..=0xDFFF).contains(&cp) })
        .collect();

    // Process id in the name: two conversions running side by side used to
    // write the same three files and read back each other's subset, producing
    // a PDF built from another document's glyph set with no error anywhere.
    let scratch = std::env::temp_dir();
    // and a counter: `serve` runs conversions side by side in one process.
    static SCRATCH_SEQ: std::sync::atomic::AtomicUsize = std::sync::atomic::AtomicUsize::new(0);
    let seq = SCRATCH_SEQ.fetch_add(1, std::sync::atomic::Ordering::Relaxed);
    let stem = format!("oxi-font-{}-{}-{}", label, std::process::id(), seq);
    let subset_path = scratch.join(format!("{stem}.otf")).to_string_lossy().replace('\\', "/");
    let cidmap_path = scratch.join(format!("{stem}-cidmap.json")).to_string_lossy().replace('\\', "/");
    let widths_path = scratch.join(format!("{stem}-widths.json")).to_string_lossy().replace('\\', "/");
    let font_path_escaped = font_path.replace('\\', "\\\\");
    let python_script = format!(
        r#"
import sys, json
sys.stdin.reconfigure(encoding='utf-8', errors='replace')
from fontTools.ttLib import TTCollection, TTFont
from fontTools.subset import Subsetter, Options

font_path = "{font_path_escaped}"
data = open(font_path, 'rb').read()

if data[:4] == b'ttcf':
    ttc = TTCollection(font_path)
    font = ttc[0]
else:
    font = TTFont(font_path)

chars = sys.stdin.read()
chars = ''.join(c for c in chars if ord(c) < 0xD800 or ord(c) > 0xDFFF)

# Read the PostScript name BEFORE subsetting. Apple ships collections whose
# nameID 6 lives only on platform 1, and the subsetter's defaults
# (name_legacy=False, name_languages=[1033]) prune exactly that record -- so
# reading it afterwards returned nothing and the PDF went out carrying our
# internal handle as its BaseFont. Reading first needs no option tuning and
# survives whatever a future fontTools decides to drop.
ps_name = None
for record in font['name'].names:
    if record.nameID == 6:
        try:
            ps_name = record.toUnicode()
            break
        except Exception:
            pass

opts = Options()
subsetter = Subsetter(options=opts)
subsetter.populate(text=chars)
subsetter.subset(font)

font.save("{subset_path}")

# A subset that kept no glyph has no cmap left, and getBestCmap() returns
# None rather than an empty map. That happens whenever none of the requested
# characters exist in this face -- a symbol font asked for Latin, say -- and
# it used to raise on the next line and take the whole face down.
try:
    cmap = font.getBestCmap()
except Exception:
    cmap = None
if cmap is None:
    cmap = {{}}
unicode_to_cid = {{}}
cid_widths = {{}}

upem = font['head'].unitsPerEm
hmtx = font['hmtx']

for unicode_val, glyph_name in cmap.items():
    cid = None
    if glyph_name.startswith('cid'):
        cid = int(glyph_name[3:])
    elif glyph_name == '.notdef':
        continue
    else:
        glyph_order = font.getGlyphOrder()
        try:
            cid = glyph_order.index(glyph_name)
        except ValueError:
            continue

    if cid is not None:
        unicode_to_cid[unicode_val] = cid
        if glyph_name in hmtx.metrics:
            advance, _ = hmtx.metrics[glyph_name]
            w = int(advance * 1000 / upem)
            cid_widths[cid] = w

with open("{cidmap_path}", 'w') as f:
    json.dump(unicode_to_cid, f)

with open("{widths_path}", 'w') as f:
    json.dump(cid_widths, f)

if ps_name:
    with open("{widths_path}.psname", 'w') as f:
        f.write(ps_name)

print(f"OK {{len(unicode_to_cid)}} mappings, {{len(cid_widths)}} widths")
"#
    );

    let result = std::process::Command::new(PYTHON3)
        .arg("-c")
        .arg(&python_script)
        .stdin(std::process::Stdio::piped())
        .stdout(std::process::Stdio::piped())
        .stderr(std::process::Stdio::piped())
        .spawn();

    match result {
        Ok(mut child) => {
            if let Some(ref mut stdin) = child.stdin {
                use std::io::Write;
                let _ = stdin.write_all(char_string.as_bytes());
            }
            let output = child.wait_with_output().ok()?;
            if !output.status.success() {
                eprintln!("Font subsetting ({}) failed: {}", label, String::from_utf8_lossy(&output.stderr));
                clear_scratch(&widths_path, &cidmap_path, &subset_path);
                return embed_whole_font(font_path, label);
            }
            note!("Font subset ({}): {}", label, String::from_utf8_lossy(&output.stdout).trim());
        }
        Err(e) => {
            eprintln!("Python not available for {} subsetting: {}", label, e);
            clear_scratch(&widths_path, &cidmap_path, &subset_path);
            return embed_whole_font(font_path, label);
        }
    }

    let otf_data = match fs::read(&subset_path) {
        Ok(data) => data,
        Err(_) => {
            clear_scratch(&widths_path, &cidmap_path, &subset_path);
            return embed_whole_font(font_path, label);
        }
    };
    let unicode_to_gid = {
        let from_subsetter = fs::read_to_string(&cidmap_path)
            .map(|json| parse_cidmap_json(&json))
            .unwrap_or_default();
        if from_subsetter.is_empty() {
            // The subsetter reports through `getBestCmap()`, which knows only
            // Unicode cmaps and answers None for a symbol font — Wingdings and
            // Symbol carry a (3,0) table and nothing else. Read the subset
            // ourselves instead of trusting the report.
            parse_cmap_table(&otf_data)
        } else {
            from_subsetter
        }
    };

    if unicode_to_gid.is_empty() {
        // Still nothing: the subset kept no glyph we can address, so embedding
        // it would put a font in the file that draws nothing — a PDF that
        // passes every structural check and shows blank where the bullets
        // were. The whole font has the mapping the subset lost.
        note!(
            "Font subset ({}) mapped no glyph; embedding the whole font instead",
            label
        );
        clear_scratch(&widths_path, &cidmap_path, &subset_path);
        return embed_whole_font(font_path, label);
    }
    let cid_widths = if let Ok(json_str) = fs::read_to_string(&widths_path) {
        parse_cidwidths_json(&json_str)
    } else {
        HashMap::new()
    };

    // Read PostScript name from font subsetting output
    let ps_name = fs::read_to_string(format!("{}.psname", &widths_path)).ok()
        .map(|s| s.trim().to_string())
        .filter(|s| !s.is_empty());
    if let Some(ref psn) = ps_name {
        note!("PostScript name: {}", psn);
    }

    clear_scratch(&widths_path, &cidmap_path, &subset_path);

    let is_cff = otf_data.starts_with(b"OTTO") || has_cff_table(&otf_data);
    if is_cff {
        if let Some(cff) = extract_cff_from_otf(&otf_data) {
            return Some(EmbeddedFont { ps_name: ps_name.clone(),
                data: cff,
                format: FontFormat::OpenTypeCff,
                unicode_to_gid,
                cid_widths,
            });
        }
    }

    Some(EmbeddedFont { ps_name,
        data: otf_data,
        format: FontFormat::TrueType,
        unicode_to_gid,
        cid_widths,
    })
}

fn collect_chars_from_block(block: &Block, chars: &mut std::collections::BTreeSet<char>) {
    match block {
        Block::Paragraph(para) => {
            for run in &para.runs {
                for ch in run.text.chars() {
                    chars.insert(ch);
                }
            }
            if let Some(ref marker) = para.style.list_marker {
                for ch in marker.chars() {
                    chars.insert(ch);
                }
            }
        }
        Block::Table(table) => {
            for row in &table.rows {
                for cell in &row.cells {
                    for block in &cell.blocks {
                        collect_chars_from_block(block, chars);
                    }
                }
            }
        }
        _ => {}
    }
}

fn collect_cjk_fonts_from_block(block: &Block, needs_cjk: &mut bool, font_names: &mut Vec<String>) {
    match block {
        Block::Paragraph(para) => {
            for run in &para.runs {
                if run.text.chars().any(is_cjk_char) {
                    *needs_cjk = true;
                }
                if let Some(ref family) = run.style.font_family {
                    if !font_names.contains(family) {
                        font_names.push(family.clone());
                    }
                }
                if let Some(ref family) = run.style.font_family_east_asia {
                    if !font_names.contains(family) {
                        font_names.push(family.clone());
                    }
                }
            }
            if let Some(ref drs) = para.style.default_run_style {
                if let Some(ref family) = drs.font_family {
                    if !font_names.contains(family) {
                        font_names.push(family.clone());
                    }
                }
            }
        }
        Block::Table(table) => {
            for row in &table.rows {
                for cell in &row.cells {
                    for block in &cell.blocks {
                        collect_cjk_fonts_from_block(block, needs_cjk, font_names);
                    }
                }
            }
        }
        _ => {}
    }
}

fn find_and_subset_cjk_font(used_chars: &std::collections::BTreeSet<char>) -> Option<EmbeddedFont> {
    // The shipped Gothic is the last resort, so a machine with no Japanese
    // face at all -- a server container, typically -- still draws the text
    // rather than warning and dropping it.
    let font_path = CJK_FONT_PATHS
        .iter()
        .map(std::path::PathBuf::from)
        .find(|p| p.exists())
        .or_else(|| bundled_face_file("OxiGothic.ttf"))?;
    note!("Found CJK font: {}", font_path.display());
    subset_font_file(&font_path.to_string_lossy(), used_chars, "cjk")
}

fn find_and_subset_cjk_font_bold(used_chars: &std::collections::BTreeSet<char>) -> Option<EmbeddedFont> {
    let font_path = CJK_FONT_PATHS_BOLD.iter().find(|p| std::path::Path::new(p).exists())?;
    note!("Found CJK Bold font: {}", font_path);
    // Was a 144-line copy of the subsetter with its own hardcoded scratch
    // paths and no way out when the subsetter was missing, so bold CJK text
    // simply vanished on a machine without fontTools -- and on Windows the
    // "/tmp/..." it wrote to means C:	mp\, which need not exist at all.
    subset_font_file(font_path, used_chars, "cjk-bold")
}

/// Fallback: load the full font without subsetting.
fn find_system_cjk_font_full() -> Option<EmbeddedFont> {
    find_system_cjk_font()
}

/// Find a CJK font on the system and extract it for PDF embedding (no subsetting).
fn find_system_cjk_font() -> Option<EmbeddedFont> {
    for path in CJK_FONT_PATHS {
        if let Ok(data) = fs::read(path) {
            note!("Found CJK font: {}", path);

            // Get the standalone font data (extract from TTC if needed)
            let otf_data = if data.len() > 12 && &data[0..4] == b"ttcf" {
                extract_font_from_ttc(&data, 0)?
            } else {
                data
            };

            // Parse cmap table to get Unicode → GlyphID mapping
            let unicode_to_gid = parse_cmap_table(&otf_data);
            note!("Parsed cmap: {} Unicode→GID entries", unicode_to_gid.len());

            // Determine format and extract appropriate data for embedding
            let is_cff = otf_data.starts_with(b"OTTO") || has_cff_table(&otf_data);
            if is_cff {
                if let Some(cff) = extract_cff_from_otf(&otf_data) {
                    return Some(EmbeddedFont { ps_name: None,
                        data: cff,
                        format: FontFormat::OpenTypeCff,
                        unicode_to_gid,
                        cid_widths: HashMap::new(),
                    });
                }
            }

            // TrueType
            return Some(EmbeddedFont { ps_name: None,
                data: otf_data,
                format: FontFormat::TrueType,
                unicode_to_gid,
                cid_widths: HashMap::new(),
            });
        }
    }
    None
}

/// Extract a single font from a TTC (TrueType Collection).
fn extract_font_from_ttc(data: &[u8], font_index: usize) -> Option<Vec<u8>> {
    if data.len() < 12 || &data[0..4] != b"ttcf" {
        return None;
    }
    let num_fonts = u32::from_be_bytes([data[8], data[9], data[10], data[11]]) as usize;
    if font_index >= num_fonts {
        return None;
    }

    let offset_pos = 12 + font_index * 4;
    if offset_pos + 4 > data.len() {
        return None;
    }
    let font_offset = u32::from_be_bytes([
        data[offset_pos], data[offset_pos + 1],
        data[offset_pos + 2], data[offset_pos + 3],
    ]) as usize;

    if font_offset + 12 > data.len() {
        return None;
    }

    // Read table directory at font_offset
    let num_tables = u16::from_be_bytes([data[font_offset + 4], data[font_offset + 5]]) as usize;
    let table_dir_start = font_offset + 12;

    // Collect all table records
    struct TableRecord {
        tag: [u8; 4],
        checksum: u32,
        offset: usize,
        length: usize,
    }

    let mut tables = Vec::new();
    for i in 0..num_tables {
        let rec_off = table_dir_start + i * 16;
        if rec_off + 16 > data.len() { return None; }
        let mut tag = [0u8; 4];
        tag.copy_from_slice(&data[rec_off..rec_off + 4]);
        let checksum = u32::from_be_bytes([
            data[rec_off + 4], data[rec_off + 5], data[rec_off + 6], data[rec_off + 7],
        ]);
        let offset = u32::from_be_bytes([
            data[rec_off + 8], data[rec_off + 9], data[rec_off + 10], data[rec_off + 11],
        ]) as usize;
        let length = u32::from_be_bytes([
            data[rec_off + 12], data[rec_off + 13], data[rec_off + 14], data[rec_off + 15],
        ]) as usize;
        tables.push(TableRecord { tag, checksum, offset, length });
    }

    // Build a standalone font file
    // OTF/TTF header: sfVersion(4) + numTables(2) + searchRange(2) + entrySelector(2) + rangeShift(2) = 12
    // Table directory: numTables * 16
    let header_size = 12 + num_tables * 16;
    let mut font = Vec::new();

    // Copy sfVersion from the original font
    font.extend_from_slice(&data[font_offset..font_offset + 4]);

    // numTables, searchRange, entrySelector, rangeShift
    font.extend_from_slice(&(num_tables as u16).to_be_bytes());

    // Calculate search params
    let mut search_range = 1u16;
    let mut entry_selector = 0u16;
    while search_range * 2 <= num_tables as u16 {
        search_range *= 2;
        entry_selector += 1;
    }
    search_range *= 16;
    let range_shift = (num_tables as u16) * 16 - search_range;
    font.extend_from_slice(&search_range.to_be_bytes());
    font.extend_from_slice(&entry_selector.to_be_bytes());
    font.extend_from_slice(&range_shift.to_be_bytes());

    // Calculate new offsets for each table
    let mut current_offset = header_size;
    let mut new_offsets = Vec::new();
    for t in &tables {
        new_offsets.push(current_offset);
        current_offset += (t.length + 3) & !3; // 4-byte align
    }

    // Write table directory
    for (i, t) in tables.iter().enumerate() {
        font.extend_from_slice(&t.tag);
        font.extend_from_slice(&t.checksum.to_be_bytes());
        font.extend_from_slice(&(new_offsets[i] as u32).to_be_bytes());
        font.extend_from_slice(&(t.length as u32).to_be_bytes());
    }

    // Write table data
    for t in &tables {
        if t.offset + t.length > data.len() { return None; }
        font.extend_from_slice(&data[t.offset..t.offset + t.length]);
        // Pad to 4-byte boundary
        while font.len() % 4 != 0 {
            font.push(0);
        }
    }

    Some(font)
}

/// Parse Unicode→CID mapping from JSON output by Python fonttools.
fn parse_cidmap_json(json_str: &str) -> HashMap<u32, u16> {
    let mut result = HashMap::new();
    // Simple JSON parsing: {"unicode": cid, ...}
    // Keys are stringified Unicode codepoints, values are CID numbers
    if let Ok(map) = serde_json::from_str::<HashMap<String, u64>>(json_str) {
        for (key, cid) in map {
            if let Ok(unicode) = key.parse::<u32>() {
                if cid <= 0xFFFF {
                    result.insert(unicode, cid as u16);
                }
            }
        }
    }
    result
}

/// Parse CID → width mapping from JSON output by Python fonttools.
fn parse_cidwidths_json(json_str: &str) -> HashMap<u16, u16> {
    let mut result = HashMap::new();
    if let Ok(map) = serde_json::from_str::<HashMap<String, u64>>(json_str) {
        for (key, width) in map {
            if let Ok(cid) = key.parse::<u16>() {
                if width <= 0xFFFF {
                    result.insert(cid, width as u16);
                }
            }
        }
    }
    result
}

/// Parse the cmap table from an OTF/TTF font to extract Unicode → GlyphID mappings.
fn parse_cmap_table(font_data: &[u8]) -> HashMap<u32, u16> {
    let mut result = HashMap::new();
    if font_data.len() < 12 { return result; }

    let num_tables = u16::from_be_bytes([font_data[4], font_data[5]]) as usize;

    // Find the cmap table
    let mut cmap_offset = 0usize;
    let mut cmap_length = 0usize;
    for i in 0..num_tables {
        let rec = 12 + i * 16;
        if rec + 16 > font_data.len() { break; }
        if &font_data[rec..rec + 4] == b"cmap" {
            cmap_offset = u32::from_be_bytes([
                font_data[rec + 8], font_data[rec + 9], font_data[rec + 10], font_data[rec + 11],
            ]) as usize;
            cmap_length = u32::from_be_bytes([
                font_data[rec + 12], font_data[rec + 13], font_data[rec + 14], font_data[rec + 15],
            ]) as usize;
            break;
        }
    }
    if cmap_offset == 0 || cmap_offset + 4 > font_data.len() { return result; }

    let cmap = &font_data[cmap_offset..font_data.len().min(cmap_offset + cmap_length)];
    if cmap.len() < 4 { return result; }

    let num_subtables = u16::from_be_bytes([cmap[2], cmap[3]]) as usize;

    // Prefer: Platform 3 Encoding 10 (Windows UCS-4, Format 12) > Platform 3 Encoding 1 (Windows BMP, Format 4)
    // > Platform 0 (Unicode)
    let mut best_offset = 0usize;
    let mut best_priority = 0u8;

    for i in 0..num_subtables {
        let rec = 4 + i * 8;
        if rec + 8 > cmap.len() { break; }
        let platform = u16::from_be_bytes([cmap[rec], cmap[rec + 1]]);
        let encoding = u16::from_be_bytes([cmap[rec + 2], cmap[rec + 3]]);
        let offset = u32::from_be_bytes([cmap[rec + 4], cmap[rec + 5], cmap[rec + 6], cmap[rec + 7]]) as usize;

        let priority = match (platform, encoding) {
            (3, 10) => 4, // Windows UCS-4 (best, supports all Unicode)
            (0, 4) => 3,  // Unicode full
            (3, 1) => 2,  // Windows BMP
            (0, 3) => 2,  // Unicode BMP
            (0, _) => 1,  // Any Unicode platform
            _ => 0,
        };

        if priority > best_priority {
            best_priority = priority;
            best_offset = offset;
        }
    }

    if best_offset == 0 || best_offset + 2 > cmap.len() { return result; }

    let subtable = &cmap[best_offset..];
    if subtable.len() < 2 { return result; }
    let format = u16::from_be_bytes([subtable[0], subtable[1]]);

    match format {
        4 => parse_cmap_format4(subtable, &mut result),
        12 => parse_cmap_format12(subtable, &mut result),
        _ => {
            eprintln!("Unsupported cmap format: {}", format);
        }
    }

    result
}

fn parse_cmap_format4(data: &[u8], result: &mut HashMap<u32, u16>) {
    if data.len() < 14 { return; }
    let seg_count = u16::from_be_bytes([data[6], data[7]]) as usize / 2;
    let header_size = 14;

    if data.len() < header_size + seg_count * 8 { return; }

    let end_codes = &data[header_size..];
    let start_codes = &data[header_size + seg_count * 2 + 2..]; // +2 for reservedPad
    let id_deltas = &data[header_size + seg_count * 4 + 2..];
    let id_range_offsets_start = header_size + seg_count * 6 + 2;
    let id_range_offsets = &data[id_range_offsets_start..];

    for seg in 0..seg_count {
        let off = seg * 2;
        if off + 2 > end_codes.len() || off + 2 > start_codes.len() { break; }
        let end_code = u16::from_be_bytes([end_codes[off], end_codes[off + 1]]);
        let start_code = u16::from_be_bytes([start_codes[off], start_codes[off + 1]]);
        if off + 2 > id_deltas.len() || off + 2 > id_range_offsets.len() { break; }
        let id_delta = i16::from_be_bytes([id_deltas[off], id_deltas[off + 1]]);
        let id_range_offset = u16::from_be_bytes([id_range_offsets[off], id_range_offsets[off + 1]]);

        if start_code == 0xFFFF { break; }

        for code in start_code..=end_code {
            let gid = if id_range_offset == 0 {
                (code as i32 + id_delta as i32) as u16
            } else {
                // idRangeOffset points into the glyphIdArray
                let glyph_idx_offset = id_range_offsets_start + off
                    + id_range_offset as usize
                    + (code - start_code) as usize * 2;
                if glyph_idx_offset + 2 <= data.len() {
                    let glyph_id = u16::from_be_bytes([data[glyph_idx_offset], data[glyph_idx_offset + 1]]);
                    if glyph_id == 0 {
                        0
                    } else {
                        (glyph_id as i32 + id_delta as i32) as u16
                    }
                } else {
                    0
                }
            };
            if gid != 0 {
                result.insert(code as u32, gid);
            }
        }
    }
}

fn parse_cmap_format12(data: &[u8], result: &mut HashMap<u32, u16>) {
    if data.len() < 16 { return; }
    let num_groups = u32::from_be_bytes([data[12], data[13], data[14], data[15]]) as usize;

    for i in 0..num_groups {
        let off = 16 + i * 12;
        if off + 12 > data.len() { break; }
        let start_code = u32::from_be_bytes([data[off], data[off + 1], data[off + 2], data[off + 3]]);
        let end_code = u32::from_be_bytes([data[off + 4], data[off + 5], data[off + 6], data[off + 7]]);
        let start_gid = u32::from_be_bytes([data[off + 8], data[off + 9], data[off + 10], data[off + 11]]);

        for code in start_code..=end_code {
            let gid = start_gid + (code - start_code);
            if gid != 0 && gid <= 0xFFFF {
                result.insert(code, gid as u16);
            }
        }
    }
}

fn has_cff_table(data: &[u8]) -> bool {
    if data.len() < 12 { return false; }
    let num_tables = u16::from_be_bytes([data[4], data[5]]) as usize;
    for i in 0..num_tables {
        let off = 12 + i * 16;
        if off + 4 > data.len() { return false; }
        if &data[off..off + 4] == b"CFF " {
            return true;
        }
    }
    false
}

/// Extract raw CFF data from an OTF font file.
fn extract_cff_from_otf(data: &[u8]) -> Option<Vec<u8>> {
    if data.len() < 12 { return None; }
    let num_tables = u16::from_be_bytes([data[4], data[5]]) as usize;
    for i in 0..num_tables {
        let rec_off = 12 + i * 16;
        if rec_off + 16 > data.len() { return None; }
        if &data[rec_off..rec_off + 4] == b"CFF " {
            let offset = u32::from_be_bytes([
                data[rec_off + 8], data[rec_off + 9], data[rec_off + 10], data[rec_off + 11],
            ]) as usize;
            let length = u32::from_be_bytes([
                data[rec_off + 12], data[rec_off + 13], data[rec_off + 14], data[rec_off + 15],
            ]) as usize;
            if offset + length <= data.len() {
                return Some(data[offset..offset + length].to_vec());
            }
        }
    }
    None
}

#[cfg(test)]
mod serve_tests {
    use super::*;
    use serde_json::json;

    /// A bad request, a missing file and a document the strict check refuses
    /// each answer with the status `docx-to-pdf` would exit with, and carry
    /// their id back.
    #[test]
    fn requests_answer_with_docx_to_pdf_statuses() {
        let bad = serve_one("not json");
        assert_eq!((bad["ok"].as_bool(), bad["exit"].as_i64()), (Some(false), Some(EXIT_FAILED as i64)));

        let missing = serve_one(r#"{"id": 7, "input": "/no/such.docx", "output": "/tmp/x.pdf"}"#);
        assert_eq!(missing["id"], 7);
        assert_eq!(missing["exit"].as_i64(), Some(EXIT_FAILED as i64));

        let no_output = serve_one(r#"{"id": "a", "input": "x.docx"}"#);
        assert_eq!(no_output["id"], "a");
        assert_eq!(no_output["ok"].as_bool(), Some(false));
    }

    /// An output ending in .png draws the asked page fitted inside the box,
    /// and a page past the end fails like a bad document.
    #[test]
    fn a_png_request_draws_one_page_inside_the_box() {
        let input = concat!(env!("CARGO_MANIFEST_DIR"), "/../../tests/fixtures/table_default_docx.docx");
        let dir = std::env::temp_dir().join(format!("oxi-png-{}", std::process::id()));
        fs::create_dir_all(&dir).unwrap();
        let output = dir.join("thumb.PNG");
        let request = json!({ "id": 1, "input": input, "output": output, "width": 200, "height": 200 });
        let answer = serve_one(&request.to_string());
        assert_eq!(answer["ok"].as_bool(), Some(true), "{answer}");
        let png = image::open(&output).unwrap();
        assert!(png.width() <= 200 && png.height() == 200, "{}x{}", png.width(), png.height());

        let past_end = json!({ "input": input, "output": output, "page": 99 });
        assert_eq!(serve_one(&past_end.to_string())["exit"].as_i64(), Some(EXIT_FAILED as i64));
        let _ = fs::remove_dir_all(&dir);
    }

    /// One conversion's font plan must not reach the next one on the same
    /// thread: serve converts many documents per process, and a stale plan
    /// drew every later document's runs with the first one's faces.
    #[test]
    fn a_conversion_starts_from_an_empty_font_plan() {
        FONT_PLAN.with(|p| {
            p.borrow_mut().keys.insert(("arial".to_string(), false, false), "FStale".to_string());
        });
        assert_eq!(resolve_font_styled(Some("Arial"), false, false), "FStale");
        let empty = oxidocs_core::layout::LayoutResult { pages: Vec::new(), hf_parts_shown: Default::default() };
        embed_document_faces(&empty, &mut HashMap::new());
        assert_ne!(resolve_font_styled(Some("Arial"), false, false), "FStale");
    }
}

#[cfg(test)]
mod coverage_tests {
    use super::*;

    #[test]
    fn story_text_leaves_out_results_the_engine_recomputes() {
        let xml = r#"<w:hdr><w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> STYLEREF  Divider </w:instrText></w:r>
            <w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>Error! No text of specified style in document.</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>
            <w:r><w:t>Table 1</w:t></w:r><w:fldSimple w:instr=" PAGE "><w:r><w:t>7</w:t></w:r></w:fldSimple>
            <w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText>REF note</w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>
            <w:r><w:t>kept</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r><w:fldSimple w:instr="DATE"><w:r><w:t>today</w:t></w:r></w:fldSimple></w:p></w:hdr>"#;
        assert_eq!(story_text(xml, &Default::default()), vec!["Table 1", "kept", "today"]);
    }

    #[test]
    fn story_text_keeps_visible_runs_once() {
        let xml = r#"<w:body><w:p><w:r><w:rPr><w:rFonts w:ascii="A"/></w:rPr><w:t>One</w:t></w:r>
            <w:r><w:rPr><w:vanish/></w:rPr><w:t>hidden</w:t></w:r>
            <w:r><w:rPr><w:vanish w:val="0"/></w:rPr><w:t xml:space="preserve"> two &amp; &#x41;</w:t></w:r>
            <w:r><mc:AlternateContent><mc:Choice><w:txbxContent><w:p><w:r><w:t>box</w:t></w:r></w:p></w:txbxContent></mc:Choice>
            <mc:Fallback><w:txbxContent><w:p><w:r><w:t>box</w:t></w:r></w:p></w:txbxContent></mc:Fallback></mc:AlternateContent></w:r>
            <w:r><w:t/><w:tab/><w:t>end</w:t></w:r></w:p></w:body>"#;
        assert_eq!(story_text(xml, &Default::default()), vec!["One", " two & A", "box", "end"]);
    }

    #[test]
    fn story_text_leaves_out_what_styles_hide() {
        let styles = r#"<w:styles><w:style w:type="paragraph" w:styleId="Note"><w:rPr><w:vanish/></w:rPr></w:style>
            <w:style w:type="paragraph" w:styleId="NoteChild"><w:basedOn w:val="Note"/></w:style>
            <w:style w:type="paragraph" w:styleId="Shown"><w:basedOn w:val="Note"/><w:rPr><w:vanish w:val="0"/></w:rPr></w:style>
            <w:style w:type="character" w:styleId="Secret"><w:rPr><w:vanish/></w:rPr></w:style></w:styles>"#;
        let hidden = hidden_styles(styles);
        let xml = r#"<w:body><w:p><w:pPr><w:pStyle w:val="NoteChild"/></w:pPr><w:r><w:t>note</w:t></w:r>
            <w:r><w:rPr><w:vanish w:val="0"/></w:rPr><w:t>shown anyway</w:t></w:r></w:p>
            <w:p><w:pPr><w:pStyle w:val="Shown"/></w:pPr><w:r><w:t>body</w:t></w:r><w:r><w:rPr><w:rStyle w:val="Secret"/></w:rPr><w:t>secret</w:t></w:r></w:p>
            <w:p><w:r><w:t>plain</w:t></w:r></w:p></w:body>"#;
        assert_eq!(story_text(xml, &hidden), vec!["shown anyway", "body", "plain"]);
    }

    #[test]
    fn toggle_in_reads_on_off_elements() {
        assert_eq!(toggle_in("<w:sectPr><w:titlePg/></w:sectPr>", "<w:titlePg"), Some(true));
        assert_eq!(toggle_in(r#"<w:titlePg w:val="0"/>"#, "<w:titlePg"), Some(false));
        assert_eq!(toggle_in("<w:titlePgX/>", "<w:titlePg"), None);
        assert_eq!(toggle_in("<w:specVanish/>", "<w:vanish"), None);
    }

    #[test]
    fn unescape_xml_leaves_unknown_entities() {
        assert_eq!(unescape_xml("a&lt;b&#65;&nbsp;&"), "a<bA&nbsp;&");
    }
}
