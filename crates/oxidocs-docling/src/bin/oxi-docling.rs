// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! oxi-docling annotate <input.docx> <docling.json> [output.json]
//!     Fill prov/pages of the DoclingDocument docling produced for <input.docx>.
//! oxi-docling <input.docx> [output.json]
//!     Standalone: Oxi's own DoclingDocument, for use without docling.
//! Output goes to stdout when no output path is given.
fn die(msg: impl std::fmt::Display) -> ! {
    eprintln!("{msg}");
    std::process::exit(1)
}

fn main() {
    let args: Vec<String> = std::env::args().skip(1).collect();
    let usage = "usage: oxi-docling annotate <input.docx> <docling.json> [output.json]\n       oxi-docling <input.docx> [output.json]";
    let (out, path) = if args.first().map(String::as_str) == Some("annotate") {
        let (Some(docx), Some(json)) = (args.get(1), args.get(2)) else { die(usage) };
        let data = std::fs::read(docx).unwrap_or_else(|e| die(format!("{docx}: {e}")));
        let raw = std::fs::read(json).unwrap_or_else(|e| die(format!("{json}: {e}")));
        let d: serde_json::Value = serde_json::from_slice(&raw).unwrap_or_else(|e| die(format!("{json}: {e}")));
        let (out, s) = oxidocs_docling::annotate::annotate(&data, d).unwrap_or_else(|e| die(e));
        eprintln!("{docx}: {} pages, {} items, {} with provenance", s.pages, s.items, s.items_with_prov);
        for (r, t) in &s.unmatched {
            eprintln!("  unplaced {r} {t:?}");
        }
        (out, args.get(3))
    } else {
        let Some(input) = args.first() else { die(usage) };
        let data = std::fs::read(input).unwrap_or_else(|e| die(format!("{input}: {e}")));
        let (out, s) = oxidocs_docling::docx_to_docling(&data, input).unwrap_or_else(|e| die(e));
        eprintln!("{input}: {} pages, {} items, {} with provenance, {} unplaced fragments",
            s.pages, s.items, s.items_with_prov, s.unplaced_fragments);
        (out, args.get(1))
    };
    let text = serde_json::to_string_pretty(&out).unwrap();
    match path {
        Some(p) => std::fs::write(p, text).unwrap_or_else(|e| die(format!("{p}: {e}"))),
        None => println!("{text}"),
    }
}
